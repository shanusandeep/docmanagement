using Microsoft.Extensions.Options;

namespace Scenario1.TransientEditing.Services;

/// <summary>
/// The scheduled backstop job (design doc, Section 5, step 5).
/// Sweeps all checked-out documents and syncs back the idle, unlocked ones.
/// </summary>
public class SyncBackWorker : BackgroundService
{
    private readonly AppOnlyGraphProvider _appOnly;
    private readonly DocumentStateStore _store;
    private readonly SyncBackEngine _engine;
    private readonly ActivityLog _log;
    private readonly DemoOptions _demo;

    public SyncBackWorker(
        AppOnlyGraphProvider appOnly,
        DocumentStateStore store,
        SyncBackEngine engine,
        ActivityLog log,
        IOptions<DemoOptions> demo)
    {
        _appOnly = appOnly;
        _store = store;
        _engine = engine;
        _log = log;
        _demo = demo.Value;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        if (!_appOnly.IsConfigured)
        {
            _log.Warn("App-only Graph credentials not configured — the automatic sync-back sweep is disabled. Use the manual 'Sync back now' button (runs as the signed-in user).");
            return;
        }

        _log.Info($"Sync-back sweep running every {_demo.SweepIntervalSeconds}s; idle threshold {_demo.IdleThresholdSeconds}s.");

        using var timer = new PeriodicTimer(TimeSpan.FromSeconds(_demo.SweepIntervalSeconds));
        while (await timer.WaitForNextTickAsync(stoppingToken))
        {
            foreach (var state in _store.All())
            {
                try
                {
                    await _engine.TrySyncBackAsync(_appOnly.Client!, state, force: false);
                }
                catch (Exception ex)
                {
                    _log.Error($"Sweep error on '{state.FileName}': {ex.Message}");
                }
            }
        }
    }
}
