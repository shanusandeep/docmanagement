using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Models.ODataErrors;

namespace Scenario1.TransientEditing.Services;

/// <summary>
/// The "inferred close" machinery (design doc, Section 5):
/// inactivity detection → checkout lock probe → download → write to the
/// network share → delete from SharePoint. Works with either the delegated
/// client (manual button) or the app-only client (background sweep).
/// </summary>
public class SyncBackEngine
{
    private readonly DocumentStateStore _store;
    private readonly NetworkShareService _share;
    private readonly ActivityLog _log;
    private readonly SharePointOptions _sp;
    private readonly DemoOptions _demo;

    public SyncBackEngine(
        DocumentStateStore store,
        NetworkShareService share,
        ActivityLog log,
        IOptions<SharePointOptions> sp,
        IOptions<DemoOptions> demo)
    {
        _store = store;
        _share = share;
        _log = log;
        _sp = sp.Value;
        _demo = demo.Value;
    }

    public async Task<SyncResult> TrySyncBackAsync(GraphServiceClient graph, DocumentState state, bool force)
    {
        try
        {
            var item = await graph.Drives[_sp.DriveId].Items[state.DriveItemId].GetAsync();
            if (item == null)
            {
                _store.Remove(state.FileName);
                return SyncResult.NotFound;
            }

            // In production this timestamp is pushed to us by Graph change
            // notifications; the demo polls it, which is functionally identical.
            var lastActivity = item.LastModifiedDateTime?.UtcDateTime ?? state.LastActivityUtc;
            state.LastActivityUtc = lastActivity;
            state.LastModifiedBy = item.LastModifiedBy?.User?.DisplayName ?? state.LastModifiedBy;
            _store.Upsert(state);

            var idle = DateTime.UtcNow - lastActivity;
            if (!force && idle < TimeSpan.FromSeconds(_demo.IdleThresholdSeconds))
                return SyncResult.StillActive;

            // Lock probe: SharePoint refuses checkout while anyone has the
            // document open, so we never pull a file out from under an editor.
            try
            {
                await graph.Drives[_sp.DriveId].Items[state.DriveItemId].Checkout.PostAsync();
            }
            catch (ODataError)
            {
                _log.Info($"'{state.FileName}' is idle but still open in Word/Office Online (lock probe refused) — leaving it in SharePoint.");
                return SyncResult.Locked;
            }

            _log.Info($"'{state.FileName}' idle for {(int)idle.TotalSeconds}s and unlocked — syncing back to the network share…");

            var content = await graph.Drives[_sp.DriveId].Items[state.DriveItemId].Content.GetAsync();
            if (content == null) return SyncResult.Error;
            await _share.WriteAsync(state.FileName, content);

            if (_demo.UsePermanentDelete)
            {
                try
                {
                    await graph.Drives[_sp.DriveId].Items[state.DriveItemId].PermanentDelete.PostAsync();
                }
                catch (ODataError)
                {
                    // permanentDelete can be blocked by retention policy — fall back
                    await graph.Drives[_sp.DriveId].Items[state.DriveItemId].DeleteAsync();
                }
            }
            else
            {
                await graph.Drives[_sp.DriveId].Items[state.DriveItemId].DeleteAsync();
            }

            _store.Remove(state.FileName);
            _log.Info($"'{state.FileName}' is back on the network share (with all tracked changes) and removed from SharePoint.");
            return SyncResult.Synced;
        }
        catch (ODataError ex) when (ex.ResponseStatusCode == 404)
        {
            _store.Remove(state.FileName);
            return SyncResult.NotFound;
        }
        catch (Exception ex)
        {
            _log.Error($"Sync-back failed for '{state.FileName}': {ex.Message}");
            return SyncResult.Error;
        }
    }
}
