namespace Scenario1.TransientEditing.Services;

public class SharePointOptions
{
    public string SiteId { get; set; } = "";
    public string DriveId { get; set; } = "";
}

public class DemoOptions
{
    public string NetworkSharePath { get; set; } = "";
    public int IdleThresholdSeconds { get; set; } = 120;
    public int SweepIntervalSeconds { get; set; } = 30;
    public bool UsePermanentDelete { get; set; } = false;
}

/// <summary>A document currently checked out to SharePoint for editing.</summary>
public class DocumentState
{
    public string FileName { get; set; } = "";
    public string DriveItemId { get; set; } = "";
    public string WebUrl { get; set; } = "";
    /// <summary>Direct file path — required for ms-word: links (WebUrl is the Doc.aspx viewer).</summary>
    public string WebDavUrl { get; set; } = "";
    public string CheckedOutBy { get; set; } = "";
    public DateTime CheckedOutAtUtc { get; set; }
    public DateTime LastActivityUtc { get; set; }
    public string? LastModifiedBy { get; set; }
}

/// <summary>A file or folder on the (simulated) network share.</summary>
public record ShareEntry(string Name, string RelativePath, bool IsFolder, long? SizeBytes, DateTime ModifiedUtc, int? ChildCount);

public enum SyncResult { StillActive, Locked, Synced, NotFound, Error }
