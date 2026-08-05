using Azure.Identity;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;

namespace Scenario2.SharePointMaster.Services;

/// <summary>
/// All operations against the sealed document library, performed as the
/// app-only service principal (the library's custodian). Users hold no
/// standing access — they receive just-in-time per-document grants here,
/// and always edit under their own identity via those grants.
/// </summary>
public class SealedLibraryService
{
    private readonly GraphServiceClient? _graph;
    private readonly DocumentRegistry _registry;
    private readonly WordTemplateService _word;
    private readonly ActivityLog _log;
    private readonly SharePointOptions _sp;

    public bool IsConfigured => _graph != null && !_sp.DriveId.StartsWith("YOUR_");

    public SealedLibraryService(
        IConfiguration config,
        DocumentRegistry registry,
        WordTemplateService word,
        ActivityLog log,
        IOptions<SharePointOptions> sp)
    {
        _registry = registry;
        _word = word;
        _log = log;
        _sp = sp.Value;

        var tenantId = config["AzureAd:TenantId"];
        var clientId = config["AzureAd:ClientId"];
        var clientSecret = config["AzureAd:ClientSecret"];

        if (!string.IsNullOrWhiteSpace(tenantId) && !tenantId.StartsWith("YOUR_") &&
            !string.IsNullOrWhiteSpace(clientId) && !clientId.StartsWith("YOUR_") &&
            !string.IsNullOrWhiteSpace(clientSecret) && !clientSecret.StartsWith("set-via"))
        {
            var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
            _graph = new GraphServiceClient(credential, ["https://graph.microsoft.com/.default"]);
        }
    }

    private GraphServiceClient Graph =>
        _graph ?? throw new InvalidOperationException(
            "App-only Graph credentials are not configured (AzureAd:TenantId/ClientId/ClientSecret). Scenario 2 requires them — the service principal is the library custodian.");

    // ---- Document lifecycle -------------------------------------------------

    public async Task<RegisteredDocument> CreateDocumentAsync(string name, string caseNumber, string createdBy, string folderPath = "")
    {
        if (!name.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)) name += ".docx";

        _log.Info($"Creating '{name}' in the sealed library (as the service principal; Track Changes enforced)…");

        var bytes = _word.CreateBlankDocx($"{caseNumber} — {Path.GetFileNameWithoutExtension(name)}");
        using var ms = new MemoryStream(bytes);
        var target = string.IsNullOrEmpty(folderPath) ? name : $"{folderPath}/{name}";
        var item = await Graph.Drives[_sp.DriveId].Items["root"]
            .ItemWithPath(target).Content.PutAsync(ms)
            ?? throw new InvalidOperationException("Upload returned no driveItem.");

        var doc = new RegisteredDocument
        {
            Name = name,
            CaseNumber = caseNumber,
            DriveItemId = item.Id!,
            WebUrl = item.WebUrl ?? "",
            CreatedBy = createdBy,
            CreatedAtUtc = DateTime.UtcNow,
            LastActivityUtc = DateTime.UtcNow
        };
        _registry.AddDocument(doc);
        _log.Info($"'{name}' registered (driveItemId {item.Id}). No user has access yet — the library is sealed.");
        return doc;
    }

    /// <summary>Soft delete: the file goes to the site recycle bin (~93 days), demonstrating the retention story.</summary>
    public async Task DeleteDocumentAsync(RegisteredDocument doc)
    {
        await Graph.Drives[_sp.DriveId].Items[doc.DriveItemId].DeleteAsync();
        _registry.RemoveDocument(doc.Id);
        _log.Info($"'{doc.Name}' deleted — recoverable from the SharePoint recycle bin for ~93 days (and governed by any Purview retention policy).");
    }

    // ---- Live library listing ----------------------------------------------

    // NOTE: navigation is ID-based on purpose. ItemWithPath() returns a builder
    // whose Children property is shadowed (not overridden) — upcasting it to the
    // base DriveItemItemRequestBuilder silently lists the ROOT's children.
    /// <summary>Folders and documents straight from the SharePoint drive — no app-DB filter.</summary>
    public async Task<List<LibraryDoc>> ListLibraryAsync(string? folderId = null)
    {
        var response = await Graph.Drives[_sp.DriveId].Items[folderId ?? "root"].Children
            .GetAsync(rc => rc.QueryParameters.Orderby = ["name"]);
        return (response?.Value ?? [])
            .Select(i => new LibraryDoc(
                i.Id!, i.Name ?? "(unnamed)", i.Size,
                i.LastModifiedDateTime?.UtcDateTime,
                i.LastModifiedBy?.User?.DisplayName,
                i.WebUrl ?? "",
                IsFolder: i.Folder != null,
                ChildCount: i.Folder?.ChildCount,
                WebDavUrl: i.WebDavUrl))
            .OrderByDescending(d => d.IsFolder)
            .ThenBy(d => d.Name)
            .ToList();
    }

    public async Task<LibraryDoc> CreateFolderAsync(string? parentFolderId, string name)
    {
        var item = await Graph.Drives[_sp.DriveId].Items[parentFolderId ?? "root"].Children.PostAsync(new DriveItem
        {
            Name = name,
            Folder = new Folder(),
            AdditionalData = new Dictionary<string, object> { ["@microsoft.graph.conflictBehavior"] = "fail" }
        }) ?? throw new InvalidOperationException("Folder creation returned nothing.");
        _log.Info($"Folder '{name}' created.");
        return new LibraryDoc(item.Id!, item.Name ?? name, null, item.LastModifiedDateTime?.UtcDateTime,
            item.LastModifiedBy?.User?.DisplayName, item.WebUrl ?? "", IsFolder: true, ChildCount: 0);
    }

    public async Task RenameAsync(string itemId, string oldName, string newName)
    {
        await Graph.Drives[_sp.DriveId].Items[itemId].PatchAsync(new DriveItem { Name = newName });
        _log.Info($"'{oldName}' renamed to '{newName}'.");
    }

    /// <summary>
    /// Deletes a folder only when it is empty. Non-empty folders are refused —
    /// the caller shows the warning to the user.
    /// </summary>
    public async Task<bool> TryDeleteFolderAsync(string itemId, string name)
    {
        var item = await Graph.Drives[_sp.DriveId].Items[itemId].GetAsync();
        var count = item?.Folder?.ChildCount ?? 0;
        if (count > 0)
        {
            _log.Warn($"Folder '{name}' not deleted — it still contains {count} item(s).");
            return false;
        }
        await Graph.Drives[_sp.DriveId].Items[itemId].DeleteAsync();
        _log.Info($"Empty folder '{name}' deleted.");
        return true;
    }

    /// <summary>
    /// Uploads any file into the sealed library via a Graph upload session
    /// (uniform code path regardless of size). .docx files get Track Changes
    /// enforced before upload.
    /// </summary>
    public async Task<LibraryDoc> UploadAsync(string fileName, Stream content, string folderPath = "")
    {
        var targetPath = string.IsNullOrEmpty(folderPath) ? fileName : $"{folderPath}/{fileName}";
        _log.Info($"Uploading '{fileName}' to the sealed library…");

        using var buffer = new MemoryStream();
        await content.CopyToAsync(buffer);
        var bytes = buffer.ToArray();
        if (fileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
            bytes = _word.EnsureTrackChanges(bytes);

        var sessionBody = new Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession.CreateUploadSessionPostRequestBody
        {
            Item = new DriveItemUploadableProperties
            {
                AdditionalData = new Dictionary<string, object>
                {
                    ["@microsoft.graph.conflictBehavior"] = "rename"
                }
            }
        };
        var session = await Graph.Drives[_sp.DriveId].Items["root"]
            .ItemWithPath(targetPath).CreateUploadSession.PostAsync(sessionBody)
            ?? throw new InvalidOperationException("Could not create an upload session.");

        using var uploadStream = new MemoryStream(bytes);
        var uploadTask = new LargeFileUploadTask<DriveItem>(session, uploadStream, -1, Graph.RequestAdapter);
        var result = await uploadTask.UploadAsync();
        if (!result.UploadSucceeded || result.ItemResponse is not { } item)
            throw new InvalidOperationException("Upload did not complete.");

        _log.Info($"'{item.Name}' uploaded ({(bytes.Length + 1023) / 1024} KB){(fileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) ? " with Track Changes enforced" : "")}. No user has access until granted.");
        return new LibraryDoc(item.Id!, item.Name ?? fileName, item.Size,
            item.LastModifiedDateTime?.UtcDateTime, item.LastModifiedBy?.User?.DisplayName, item.WebUrl ?? "",
            WebDavUrl: item.WebDavUrl);
    }

    /// <summary>
    /// Replaces an existing document's content — SharePoint records it as the
    /// next version. Track Changes is enforced for .docx content.
    /// </summary>
    public async Task UploadNewVersionAsync(string itemId, string fileName, Stream content)
    {
        _log.Info($"Uploading new version of '{fileName}'…");

        using var buffer = new MemoryStream();
        await content.CopyToAsync(buffer);
        var bytes = buffer.ToArray();
        if (fileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
            bytes = _word.EnsureTrackChanges(bytes);

        var sessionBody = new Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession.CreateUploadSessionPostRequestBody
        {
            Item = new DriveItemUploadableProperties
            {
                AdditionalData = new Dictionary<string, object>
                {
                    ["@microsoft.graph.conflictBehavior"] = "replace"
                }
            }
        };
        var session = await Graph.Drives[_sp.DriveId].Items[itemId].CreateUploadSession.PostAsync(sessionBody)
            ?? throw new InvalidOperationException("Could not create an upload session.");

        using var uploadStream = new MemoryStream(bytes);
        var uploadTask = new LargeFileUploadTask<DriveItem>(session, uploadStream, -1, Graph.RequestAdapter);
        var result = await uploadTask.UploadAsync();
        if (!result.UploadSucceeded)
            throw new InvalidOperationException("Version upload did not complete.");

        _log.Info($"New version of '{fileName}' uploaded — previous versions remain restorable.");
    }

    /// <summary>Finds (or creates) the registry row for a drive item so grants can attach to it.</summary>
    public RegisteredDocument EnsureRegistered(string driveItemId, string name, string webUrl)
    {
        var doc = _registry.Documents().FirstOrDefault(d => d.DriveItemId == driveItemId);
        if (doc != null) return doc;
        doc = new RegisteredDocument
        {
            Name = name,
            CaseNumber = "-",
            DriveItemId = driveItemId,
            WebUrl = webUrl,
            CreatedBy = "(existing SharePoint document)",
            CreatedAtUtc = DateTime.UtcNow,
            LastActivityUtc = DateTime.UtcNow
        };
        _registry.AddDocument(doc);
        return doc;
    }

    public async Task DeleteItemAsync(string driveItemId, string name)
    {
        await Graph.Drives[_sp.DriveId].Items[driveItemId].DeleteAsync();
        var doc = _registry.Documents().FirstOrDefault(d => d.DriveItemId == driveItemId);
        if (doc != null) _registry.RemoveDocument(doc.Id);
        _log.Info($"'{name}' deleted — recoverable from the SharePoint recycle bin for ~93 days.");
    }

    public async Task<List<VersionInfo>> ListVersionsByItemAsync(string driveItemId)
    {
        var response = await Graph.Drives[_sp.DriveId].Items[driveItemId].Versions.GetAsync();
        var versions = response?.Value ?? [];
        return versions
            .Select((v, i) => new VersionInfo(
                v.Id ?? "", v.LastModifiedDateTime?.UtcDateTime,
                v.LastModifiedBy?.User?.DisplayName, v.Size, i == 0))
            .ToList();
    }

    public async Task RestoreVersionByItemAsync(string driveItemId, string versionId, string name)
    {
        await Graph.Drives[_sp.DriveId].Items[driveItemId].Versions[versionId].RestoreVersion.PostAsync();
        _log.Info($"Version {versionId} of '{name}' restored as the current version.");
    }

    /// <summary>Content of one specific version (opens/downloads as .docx).</summary>
    public Task<Stream?> GetVersionContentAsync(string driveItemId, string versionId) =>
        Graph.Drives[_sp.DriveId].Items[driveItemId].Versions[versionId].Content.GetAsync();

    /// <summary>Current document converted to PDF by SharePoint (Graph format=pdf).</summary>
    public Task<Stream?> GetPdfAsync(string driveItemId) =>
        Graph.Drives[_sp.DriveId].Items[driveItemId].Content
            .GetAsync(rc => rc.QueryParameters.Format = "pdf");

    // ---- Just-in-time access ------------------------------------------------

    public async Task<AccessGrant> GrantAsync(RegisteredDocument doc, string upn, string role)
    {
        var existing = _registry.FindGrant(doc.Id, upn, role);
        if (existing != null)
        {
            _log.Info($"{upn} already holds '{role}' on '{doc.Name}' — reusing the grant.");
            return existing;
        }

        _log.Info($"JIT grant: giving {upn} '{role}' on '{doc.Name}' only…");

        var body = new Microsoft.Graph.Drives.Item.Items.Item.Invite.InvitePostRequestBody
        {
            Recipients = [new DriveRecipient { Email = upn }],
            RequireSignIn = true,
            SendInvitation = false,
            Roles = [role]
        };
        var response = await Graph.Drives[_sp.DriveId].Items[doc.DriveItemId]
            .Invite.PostAsInvitePostResponseAsync(body);
        var permissionId = response?.Value?.FirstOrDefault()?.Id
            ?? throw new InvalidOperationException("Invite returned no permission.");

        var grant = new AccessGrant
        {
            DocumentId = doc.Id,
            UserUpn = upn,
            Role = role,
            PermissionId = permissionId,
            GrantedAtUtc = DateTime.UtcNow
        };
        _registry.AddGrant(grant);
        _log.Info($"Granted. Permission {permissionId} recorded in the app database; it lives exactly as long as the editing session.");
        return grant;
    }

    public async Task RevokeGrantAsync(AccessGrant grant, RegisteredDocument doc)
    {
        try
        {
            await Graph.Drives[_sp.DriveId].Items[doc.DriveItemId]
                .Permissions[grant.PermissionId].DeleteAsync();
        }
        catch (ODataError ex) when (ex.ResponseStatusCode == 404)
        {
            // already gone in SharePoint — still remove from the DB
        }
        _registry.RemoveGrant(grant.Id);
        _log.Info($"Revoked {grant.UserUpn}'s '{grant.Role}' on '{doc.Name}'. The document disappears from their SharePoint views and search.");
    }

    public async Task RevokeAllAsync(RegisteredDocument doc)
    {
        foreach (var grant in _registry.GrantsFor(doc.Id))
            await RevokeGrantAsync(grant, doc);
    }

    // ---- Inferred close -----------------------------------------------------

    public async Task<DateTime?> GetLastModifiedUtcAsync(RegisteredDocument doc)
    {
        var item = await Graph.Drives[_sp.DriveId].Items[doc.DriveItemId].GetAsync();
        return item?.LastModifiedDateTime?.UtcDateTime;
    }

    /// <summary>
    /// Lock probe: checkout fails while anyone has the document open.
    /// On success we immediately check in again — we only wanted the answer.
    /// </summary>
    public async Task<bool> IsUnlockedAsync(RegisteredDocument doc)
    {
        try
        {
            await Graph.Drives[_sp.DriveId].Items[doc.DriveItemId].Checkout.PostAsync();
            await Graph.Drives[_sp.DriveId].Items[doc.DriveItemId].Checkin.PostAsync(
                new Microsoft.Graph.Drives.Item.Items.Item.Checkin.CheckinPostRequestBody
                {
                    Comment = "JIT access probe"
                });
            return true;
        }
        catch (ODataError)
        {
            return false;
        }
    }

    // ---- Version management -------------------------------------------------

    public async Task<List<VersionInfo>> ListVersionsAsync(RegisteredDocument doc)
    {
        var response = await Graph.Drives[_sp.DriveId].Items[doc.DriveItemId].Versions.GetAsync();
        var versions = response?.Value ?? [];
        return versions
            .Select((v, i) => new VersionInfo(
                v.Id ?? "",
                v.LastModifiedDateTime?.UtcDateTime,
                v.LastModifiedBy?.User?.DisplayName,
                v.Size,
                i == 0))
            .ToList();
    }

    public async Task RestoreVersionAsync(RegisteredDocument doc, string versionId)
    {
        await Graph.Drives[_sp.DriveId].Items[doc.DriveItemId]
            .Versions[versionId].RestoreVersion.PostAsync();
        _log.Info($"Version {versionId} of '{doc.Name}' restored as the current version (native SharePoint versioning).");
    }

    // ---- Reconciliation -----------------------------------------------------

    /// <summary>
    /// Compares actual SharePoint permissions against the app database and
    /// removes any user grant the database doesn't know about. Self-healing:
    /// a missed revocation cannot survive the next sweep.
    /// </summary>
    public async Task<List<string>> ReconcileAsync(bool dryRun)
    {
        var report = new List<string>();
        var known = _registry.KnownPermissionIds();

        foreach (var doc in _registry.Documents())
        {
            var response = await Graph.Drives[_sp.DriveId].Items[doc.DriveItemId].Permissions.GetAsync();
            foreach (var perm in response?.Value ?? [])
            {
                if (perm.Id == null || known.Contains(perm.Id)) continue;
                if (perm.InheritedFrom != null) continue;                       // library-level (should not exist in a sealed library)
                if (perm.Roles?.Contains("owner") == true) continue;            // site owners / admins
                var isUserGrant = perm.GrantedToV2?.User != null ||
                                  perm.GrantedToIdentitiesV2?.Any(identity => identity.User != null) == true;
                if (!isUserGrant) continue;                                     // app/site principals

                var who = perm.GrantedToV2?.User?.DisplayName
                          ?? perm.GrantedToIdentitiesV2?.FirstOrDefault()?.User?.DisplayName
                          ?? "unknown user";
                if (dryRun)
                {
                    report.Add($"WOULD REMOVE: '{doc.Name}' — {who} ({string.Join(",", perm.Roles ?? [])}) permission {perm.Id} is not in the app database.");
                }
                else
                {
                    await Graph.Drives[_sp.DriveId].Items[doc.DriveItemId].Permissions[perm.Id].DeleteAsync();
                    report.Add($"REMOVED: '{doc.Name}' — {who} ({string.Join(",", perm.Roles ?? [])}) permission {perm.Id} was not in the app database.");
                }
            }
        }

        if (report.Count == 0)
            report.Add("Clean: every SharePoint permission matches the app database.");
        _log.Info($"Reconciliation ({(dryRun ? "preview" : "applied")}): {report.Count} finding(s).");
        return report;
    }

    // ---- Links --------------------------------------------------------------

    /// <summary>Desktop Word must get the direct file path (webDavUrl), never the Doc.aspx viewer URL.</summary>
    public static string DesktopEditLink(LibraryDoc doc) =>
        $"ms-word:ofe|u|{(string.IsNullOrEmpty(doc.WebDavUrl) ? doc.WebUrl : doc.WebDavUrl)}";

    public static string OnlineLink(LibraryDoc doc, bool edit)
    {
        var action = edit ? "action=edit" : "action=view";
        return doc.WebUrl.Contains("action=default")
            ? doc.WebUrl.Replace("action=default", action)
            : doc.WebUrl + (doc.WebUrl.Contains('?') ? "&" : "?") + "web=1&" + action;
    }
}
