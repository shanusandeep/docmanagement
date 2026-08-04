using System.Text.Json;

namespace Scenario1.TransientEditing.Services;

/// <summary>
/// Tracks which documents are currently checked out to SharePoint.
/// Persisted to a JSON file so a demo survives an app restart.
/// </summary>
public class DocumentStateStore
{
    private readonly Dictionary<string, DocumentState> _states = new(StringComparer.OrdinalIgnoreCase);
    private readonly object _lock = new();
    private readonly string _path;

    public DocumentStateStore(IWebHostEnvironment env)
    {
        var dir = Path.Combine(env.ContentRootPath, "App_Data");
        Directory.CreateDirectory(dir);
        _path = Path.Combine(dir, "checkouts.json");
        if (File.Exists(_path))
        {
            try
            {
                var loaded = JsonSerializer.Deserialize<List<DocumentState>>(File.ReadAllText(_path));
                foreach (var s in loaded ?? []) _states[s.FileName] = s;
            }
            catch { /* corrupted state file: start clean */ }
        }
    }

    public DocumentState? Get(string fileName)
    {
        lock (_lock) return _states.TryGetValue(fileName, out var s) ? s : null;
    }

    public List<DocumentState> All()
    {
        lock (_lock) return _states.Values.ToList();
    }

    public void Upsert(DocumentState state)
    {
        lock (_lock) { _states[state.FileName] = state; Save(); }
    }

    public void Remove(string fileName)
    {
        lock (_lock) { _states.Remove(fileName); Save(); }
    }

    private void Save() =>
        File.WriteAllText(_path, JsonSerializer.Serialize(_states.Values.ToList(),
            new JsonSerializerOptions { WriteIndented = true }));
}
