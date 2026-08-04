using Microsoft.Extensions.Options;

namespace Scenario2.SharePointMaster.Services;

/// <summary>
/// Automatic revocation: when a document with outstanding grants goes idle
/// and the lock probe confirms nobody has it open, all its JIT grants are
/// withdrawn. (Production adds Graph change notifications; polling is the
/// localhost-friendly equivalent.)
/// </summary>
public class RevokeWorker : BackgroundService
{
    private readonly SealedLibraryService _library;
    private readonly DocumentRegistry _registry;
    private readonly ActivityLog _log;
    private readonly DemoOptions _demo;

    public RevokeWorker(
        SealedLibraryService library,
        DocumentRegistry registry,
        ActivityLog log,
        IOptions<DemoOptions> demo)
    {
        _library = library;
        _registry = registry;
        _log = log;
        _demo = demo.Value;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        if (!_library.IsConfigured)
        {
            _log.Warn("App-only Graph credentials or DriveId not configured — automatic revocation sweep disabled.");
            return;
        }

        _log.Info($"Revocation sweep running every {_demo.SweepIntervalSeconds}s; idle threshold {_demo.IdleThresholdSeconds}s.");

        using var timer = new PeriodicTimer(TimeSpan.FromSeconds(_demo.SweepIntervalSeconds));
        while (await timer.WaitForNextTickAsync(stoppingToken))
        {
            foreach (var doc in _registry.Documents())
            {
                try
                {
                    if (_registry.GrantsFor(doc.Id).Count == 0) continue;

                    var lastModified = await _library.GetLastModifiedUtcAsync(doc) ?? doc.LastActivityUtc;
                    doc.LastActivityUtc = lastModified;
                    _registry.UpdateDocument(doc);

                    if (DateTime.UtcNow - lastModified < TimeSpan.FromSeconds(_demo.IdleThresholdSeconds))
                        continue;

                    if (!await _library.IsUnlockedAsync(doc))
                    {
                        _log.Info($"'{doc.Name}' is idle but still open somewhere (lock probe) — keeping grants for now.");
                        continue;
                    }

                    _log.Info($"'{doc.Name}' idle and unlocked — withdrawing all JIT grants.");
                    await _library.RevokeAllAsync(doc);
                }
                catch (Exception ex)
                {
                    _log.Error($"Sweep error on '{doc.Name}': {ex.Message}");
                }
            }
        }
    }
}
