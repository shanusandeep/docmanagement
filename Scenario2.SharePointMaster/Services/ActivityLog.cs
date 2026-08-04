namespace Scenario2.SharePointMaster.Services;

public record ActivityEntry(DateTime AtUtc, string Level, string Message);

/// <summary>In-memory activity feed shown live in the UI during demos.</summary>
public class ActivityLog
{
    private readonly LinkedList<ActivityEntry> _entries = new();
    private readonly object _lock = new();

    public event Action? Changed;

    public void Info(string message) => Add("info", message);
    public void Warn(string message) => Add("warn", message);
    public void Error(string message) => Add("error", message);

    private void Add(string level, string message)
    {
        lock (_lock)
        {
            _entries.AddFirst(new ActivityEntry(DateTime.UtcNow, level, message));
            while (_entries.Count > 200) _entries.RemoveLast();
        }
        Changed?.Invoke();
    }

    public IReadOnlyList<ActivityEntry> Snapshot()
    {
        lock (_lock) return _entries.ToList();
    }
}
