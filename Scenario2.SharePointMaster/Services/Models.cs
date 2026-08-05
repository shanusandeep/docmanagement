namespace Scenario2.SharePointMaster.Services;

public class SharePointOptions
{
    public string SiteId { get; set; } = "";
    public string DriveId { get; set; } = "";
}

public class DemoOptions
{
    public int IdleThresholdSeconds { get; set; } = 120;
    public int SweepIntervalSeconds { get; set; } = 30;
}

/// <summary>A document permanently stored in the sealed SharePoint library.</summary>
public class RegisteredDocument
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public string Name { get; set; } = "";
    public string CaseNumber { get; set; } = "";
    public string DriveItemId { get; set; } = "";
    public string WebUrl { get; set; } = "";
    public string CreatedBy { get; set; } = "";
    public DateTime CreatedAtUtc { get; set; }
    public DateTime LastActivityUtc { get; set; }
}

/// <summary>
/// A just-in-time permission granted on one document to one user.
/// The app database is the source of truth; SharePoint enforces it.
/// </summary>
public class AccessGrant
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public Guid DocumentId { get; set; }
    public string UserUpn { get; set; } = "";
    public string Role { get; set; } = "write";
    public string PermissionId { get; set; } = "";
    public DateTime GrantedAtUtc { get; set; }
}

public record VersionInfo(string Id, DateTime? ModifiedUtc, string? ModifiedBy, long? SizeBytes, bool IsCurrent);

/// <summary>
/// A document or folder as listed live from the SharePoint drive.
/// WebUrl is the browser (Doc.aspx viewer) link; WebDavUrl is the direct file
/// path — desktop Office protocol links (ms-word:ofe|u|…) must use WebDavUrl.
/// </summary>
public record LibraryDoc(string Id, string Name, long? SizeBytes, DateTime? ModifiedUtc, string? ModifiedBy, string WebUrl,
    bool IsFolder = false, long? ChildCount = null, string? WebDavUrl = null);
