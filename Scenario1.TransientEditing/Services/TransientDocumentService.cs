using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;

namespace Scenario1.TransientEditing.Services;

/// <summary>
/// Delegated (acts-as-the-user) operations: upload a network-share document
/// to SharePoint on first View/Edit, hand back in-place edit links, and
/// convert documents to PDF via Graph (transiently for share-resident files).
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
    /// Relative paths mirror the share's folder structure in SharePoint.
    /// </summary>
    public async Task<DocumentState> CheckoutAsync(string relativePath, string userName)
    {
        var existing = _store.Get(relativePath);
        if (existing != null)
        {
            _log.Info($"'{relativePath}' is already in SharePoint (checked out by {existing.CheckedOutBy}) — returning the existing copy so {userName} joins the same co-authoring session.");
            return existing;
        }

        _log.Info($"Uploading '{relativePath}' to SharePoint for {userName} (Track Changes enforced via settings.xml injection)…");

        var bytes = relativePath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)
            ? _word.EnsureTrackChanges(_share.Read(relativePath))
            : _share.Read(relativePath);
        using var ms = new MemoryStream(bytes);

        await EnsureSpFolderAsync(GetDir(relativePath));
        var item = await _graph.Drives[_sp.DriveId].Items["root"]
            .ItemWithPath(relativePath).Content.PutAsync(ms)
            ?? throw new InvalidOperationException("Upload returned no driveItem.");

        // webDavUrl (needed for the ms-word link) is only returned when selected
        var withDav = await _graph.Drives[_sp.DriveId].Items[item.Id!]
            .GetAsync(rc => rc.QueryParameters.Select = ["id", "webUrl", "webDavUrl"]);

        var state = new DocumentState
        {
            FileName = relativePath,
            DriveItemId = item.Id!,
            WebUrl = withDav?.WebUrl ?? item.WebUrl ?? "",
            WebDavUrl = withDav?.WebDavUrl ?? "",
            CheckedOutBy = userName,
            CheckedOutAtUtc = DateTime.UtcNow,
            LastActivityUtc = DateTime.UtcNow
        };
        _store.Upsert(state);
        _log.Info($"'{relativePath}' uploaded (driveItemId {item.Id}). Opening in Word — autosave is on; edits go straight to SharePoint.");
        return state;
    }

    /// <summary>
    /// PDF via Graph. Checked-out documents convert in place; share-resident
    /// documents are uploaded transiently, converted, and deleted again.
    /// </summary>
    public async Task<(Stream Content, string FileName)> ConvertToPdfAsync(string relativePath)
    {
        var pdfName = Path.GetFileNameWithoutExtension(relativePath) + ".pdf";
        var state = _store.Get(relativePath);
        string itemId;
        var transient = false;

        if (state != null)
        {
            itemId = state.DriveItemId;
        }
        else
        {
            var bytes = _share.Read(relativePath);
            var tmpPath = $"__pdf-tmp/{Guid.NewGuid():N}-{Path.GetFileName(relativePath)}";

            // The scratch folder is reused across conversions (it's hidden from
            // listings). Assume it exists; create it only on the first-ever run.
            DriveItem? item;
            try
            {
                using var ms = new MemoryStream(bytes);
                item = await _graph.Drives[_sp.DriveId].Items["root"]
                    .ItemWithPath(tmpPath).Content.PutAsync(ms);
            }
            catch (ODataError ex) when (ex.ResponseStatusCode == 404)
            {
                await EnsureSpFolderAsync("__pdf-tmp");
                using var ms = new MemoryStream(bytes);
                item = await _graph.Drives[_sp.DriveId].Items["root"]
                    .ItemWithPath(tmpPath).Content.PutAsync(ms);
            }

            itemId = item?.Id ?? throw new InvalidOperationException("Transient upload for PDF conversion failed.");
            transient = true;
        }

        var pdfStream = await _graph.Drives[_sp.DriveId].Items[itemId].Content
            .GetAsync(rc => rc.QueryParameters.Format = "pdf")
            ?? throw new InvalidOperationException("PDF conversion returned no content.");

        // buffer fully before deleting the transient copy
        var buffer = new MemoryStream();
        await pdfStream.CopyToAsync(buffer);
        buffer.Position = 0;

        if (transient)
        {
            // permanent delete keeps both the reused scratch folder and the recycle bin clean
            try { await _graph.Drives[_sp.DriveId].Items[itemId].PermanentDelete.PostAsync(); }
            catch
            {
                try { await _graph.Drives[_sp.DriveId].Items[itemId].DeleteAsync(); }
                catch { /* best-effort; the folder is hidden from listings either way */ }
            }
        }

        _log.Info($"'{relativePath}' converted to PDF via Graph{(transient ? " (transient copy, removed after conversion)" : "")}.");
        return (buffer, pdfName);
    }

    /// <summary>
    /// Raw download. Checked-out documents stream from SharePoint (freshest
    /// autosaved bytes); everything else comes straight from the share.
    /// </summary>
    public async Task<(Stream Content, string FileName)> DownloadAsync(string relativePath)
    {
        var fileName = Path.GetFileName(relativePath);
        var state = _store.Get(relativePath);
        if (state != null)
        {
            var stream = await _graph.Drives[_sp.DriveId].Items[state.DriveItemId].Content.GetAsync();
            if (stream != null)
            {
                var buffer = new MemoryStream();
                await stream.CopyToAsync(buffer);
                buffer.Position = 0;
                return (buffer, fileName);
            }
        }
        return (new MemoryStream(_share.Read(relativePath)), fileName);
    }

    private static string? GetDir(string relativePath)
    {
        var dir = Path.GetDirectoryName(relativePath);
        return string.IsNullOrEmpty(dir) ? null : dir.Replace('\\', '/');
    }

    /// <summary>Creates the folder chain in SharePoint so nested uploads have a parent; returns the deepest folder's id.</summary>
    private async Task<string> EnsureSpFolderAsync(string? relativeDir)
    {
        var parentId = "root";
        if (string.IsNullOrEmpty(relativeDir)) return parentId;
        foreach (var segment in relativeDir.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            try
            {
                var created = await _graph.Drives[_sp.DriveId].Items[parentId].Children.PostAsync(new DriveItem
                {
                    Name = segment,
                    Folder = new Folder(),
                    AdditionalData = new Dictionary<string, object> { ["@microsoft.graph.conflictBehavior"] = "fail" }
                });
                parentId = created!.Id!;
            }
            catch (ODataError ex) when (ex.Error?.Code?.Contains("nameAlreadyExists", StringComparison.OrdinalIgnoreCase) == true)
            {
                var children = await _graph.Drives[_sp.DriveId].Items[parentId].Children.GetAsync();
                parentId = children?.Value?.FirstOrDefault(c =>
                        c.Folder != null && string.Equals(c.Name, segment, StringComparison.OrdinalIgnoreCase))?.Id
                    ?? throw new InvalidOperationException($"Folder '{segment}' exists but could not be resolved.");
            }
        }
        return parentId;
    }

    /// <summary>Desktop Word needs the direct file path (webDavUrl), not the Doc.aspx viewer URL.</summary>
    public static string DesktopEditLink(DocumentState state) =>
        $"ms-word:ofe|u|{(string.IsNullOrEmpty(state.WebDavUrl) ? state.WebUrl : state.WebDavUrl)}";
}
