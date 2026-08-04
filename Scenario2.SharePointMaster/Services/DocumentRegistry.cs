using System.Text.Json;

namespace Scenario2.SharePointMaster.Services;

/// <summary>
/// The application database: registered documents plus the access grants
/// currently projected into SharePoint. JSON file for the demo so the
/// "access is maintained in our database" claim can be shown literally.
/// </summary>
public class DocumentRegistry
{
    public class RegistryData
    {
        public List<RegisteredDocument> Documents { get; set; } = [];
        public List<AccessGrant> Grants { get; set; } = [];
    }

    private readonly RegistryData _data;
    private readonly object _lock = new();
    private readonly string _path;

    public string FilePath => _path;

    public DocumentRegistry(IWebHostEnvironment env)
    {
        var dir = Path.Combine(env.ContentRootPath, "App_Data");
        Directory.CreateDirectory(dir);
        _path = Path.Combine(dir, "registry.json");
        if (File.Exists(_path))
        {
            try { _data = JsonSerializer.Deserialize<RegistryData>(File.ReadAllText(_path)) ?? new(); }
            catch { _data = new(); }
        }
        else
        {
            _data = new();
        }
    }

    public List<RegisteredDocument> Documents()
    {
        lock (_lock) return _data.Documents.OrderBy(d => d.CaseNumber).ThenBy(d => d.Name).ToList();
    }

    public RegisteredDocument? GetDocument(Guid id)
    {
        lock (_lock) return _data.Documents.FirstOrDefault(d => d.Id == id);
    }

    public List<AccessGrant> GrantsFor(Guid documentId)
    {
        lock (_lock) return _data.Grants.Where(g => g.DocumentId == documentId).ToList();
    }

    public List<AccessGrant> AllGrants()
    {
        lock (_lock) return _data.Grants.ToList();
    }

    public AccessGrant? FindGrant(Guid documentId, string upn, string role)
    {
        lock (_lock)
            return _data.Grants.FirstOrDefault(g =>
                g.DocumentId == documentId &&
                g.UserUpn.Equals(upn, StringComparison.OrdinalIgnoreCase) &&
                g.Role == role);
    }

    public void AddDocument(RegisteredDocument doc)
    {
        lock (_lock) { _data.Documents.Add(doc); Save(); }
    }

    public void UpdateDocument(RegisteredDocument doc)
    {
        lock (_lock) { Save(); }
    }

    public void RemoveDocument(Guid id)
    {
        lock (_lock)
        {
            _data.Documents.RemoveAll(d => d.Id == id);
            _data.Grants.RemoveAll(g => g.DocumentId == id);
            Save();
        }
    }

    public void AddGrant(AccessGrant grant)
    {
        lock (_lock) { _data.Grants.Add(grant); Save(); }
    }

    public void RemoveGrant(Guid grantId)
    {
        lock (_lock) { _data.Grants.RemoveAll(g => g.Id == grantId); Save(); }
    }

    public HashSet<string> KnownPermissionIds()
    {
        lock (_lock) return _data.Grants.Select(g => g.PermissionId).ToHashSet();
    }

    private void Save() =>
        File.WriteAllText(_path, JsonSerializer.Serialize(_data, new JsonSerializerOptions { WriteIndented = true }));
}
