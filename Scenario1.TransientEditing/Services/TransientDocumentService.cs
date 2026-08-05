using Microsoft.Extensions.Options;
using Microsoft.Graph;

namespace Scenario1.TransientEditing.Services;

/// <summary>
/// Delegated (acts-as-the-user) operations: upload a network-share document
/// to SharePoint on first View/Edit and hand back in-place edit links.
/// </summary>
public class TransientDocumentService
{
    private readonly GraphServiceClient _graph;
    private readonly DocumentStateStore _store;
    private readonly NetworkShareService _share;
    private readonly WordTemplateService _word;
    private readonly ActivityLog _log;
    private readonly SharePointOptions _sp;

    public TransientDocumentService(
        GraphServiceClient graph,
        DocumentStateStore store,
        NetworkShareService share,
        WordTemplateService word,
        ActivityLog log,
        IOptions<SharePointOptions> sp)
    {
        _graph = graph;
        _store = store;
        _share = share;
        _word = word;
        _log = log;
        _sp = sp.Value;
    }

    /// <summary>
    /// If the document is already in SharePoint, return the existing copy
    /// (the user joins the live co-authoring session). Otherwise enforce
    /// Track Changes and upload it, recording the checkout in the app DB.
    /// </summary>
    public async Task<DocumentState> CheckoutAsync(string fileName, string userName)
    {
        var existing = _store.Get(fileName);
        if (existing != null)
        {
            _log.Info($"'{fileName}' is already in SharePoint (checked out by {existing.CheckedOutBy}) — returning the existing copy so {userName} joins the same co-authoring session.");
            return existing;
        }

        _log.Info($"Uploading '{fileName}' to SharePoint for {userName} (Track Changes enforced via settings.xml injection)…");

        var bytes = _word.EnsureTrackChanges(_share.Read(fileName));
        using var ms = new MemoryStream(bytes);

        var item = await _graph.Drives[_sp.DriveId].Items["root"]
            .ItemWithPath(fileName).Content.PutAsync(ms)
            ?? throw new InvalidOperationException("Upload returned no driveItem.");

        // webDavUrl (needed for the ms-word link) is only returned when selected
        var withDav = await _graph.Drives[_sp.DriveId].Items[item.Id!]
            .GetAsync(rc => rc.QueryParameters.Select = ["id", "webUrl", "webDavUrl"]);

        var state = new DocumentState
        {
            FileName = fileName,
            DriveItemId = item.Id!,
            WebUrl = withDav?.WebUrl ?? item.WebUrl ?? "",
            WebDavUrl = withDav?.WebDavUrl ?? "",
            CheckedOutBy = userName,
            CheckedOutAtUtc = DateTime.UtcNow,
            LastActivityUtc = DateTime.UtcNow
        };
        _store.Upsert(state);
        _log.Info($"'{fileName}' uploaded (driveItemId {item.Id}). Opening in Word — autosave is on; edits go straight to SharePoint.");
        return state;
    }

    /// <summary>Desktop Word needs the direct file path (webDavUrl), not the Doc.aspx viewer URL.</summary>
    public static string DesktopEditLink(DocumentState state) =>
        $"ms-word:ofe|u|{(string.IsNullOrEmpty(state.WebDavUrl) ? state.WebUrl : state.WebDavUrl)}";

    public static string OnlineEditLink(DocumentState state) =>
        state.WebUrl.Contains("action=default")
            ? state.WebUrl.Replace("action=default", "action=edit")
            : state.WebUrl + (state.WebUrl.Contains('?') ? "&" : "?") + "web=1&action=edit";
}
