using Microsoft.Extensions.Options;

namespace Scenario1.TransientEditing.Services;

/// <summary>
/// Simulates the corporate network share (Serfis document storage), including
/// subfolders. Points at a real UNC path in production; a local folder here.
/// All paths are relative to the share root and validated against escapes.
/// </summary>
public class NetworkShareService
{
    private readonly string _root;
    private readonly WordTemplateService _word;

    public NetworkShareService(IOptions<DemoOptions> options, IWebHostEnvironment env, WordTemplateService word)
    {
        _word = word;
        _root = string.IsNullOrWhiteSpace(options.Value.NetworkSharePath)
            ? Path.Combine(env.ContentRootPath, "NetworkShare")
            : options.Value.NetworkSharePath;
        Directory.CreateDirectory(_root);
        SeedSamples();
    }

    public string RootPath => _root;

    private void SeedSamples()
    {
        if (Directory.EnumerateFileSystemEntries(_root).Any()) return;
        foreach (var name in new[] { "Investigation Report", "Interview Notes", "Evidence Summary" })
        {
            File.WriteAllBytes(Path.Combine(_root, name + ".docx"),
                _word.CreateBlankDocx($"{name} (created by Serfis on the network share)"));
        }
    }

    private string Full(string relativePath)
    {
        var combined = Path.GetFullPath(Path.Combine(_root, relativePath ?? ""));
        if (!combined.StartsWith(Path.GetFullPath(_root), StringComparison.Ordinal))
            throw new InvalidOperationException("Invalid path.");
        return combined;
    }

    private static string Join(string dir, string name) =>
        string.IsNullOrEmpty(dir) ? name : $"{dir}/{name}";

    public List<ShareEntry> List(string relativeDir = "")
    {
        var dir = Full(relativeDir);
        var entries = new List<ShareEntry>();
        foreach (var d in Directory.EnumerateDirectories(dir))
        {
            var di = new DirectoryInfo(d);
            entries.Add(new ShareEntry(di.Name, Join(relativeDir, di.Name), true, null,
                di.LastWriteTimeUtc, Directory.EnumerateFileSystemEntries(d).Count()));
        }
        foreach (var f in Directory.EnumerateFiles(dir))
        {
            var fi = new FileInfo(f);
            if (fi.Name.StartsWith('.')) continue;
            entries.Add(new ShareEntry(fi.Name, Join(relativeDir, fi.Name), false, fi.Length,
                fi.LastWriteTimeUtc, null));
        }
        return entries.OrderByDescending(e => e.IsFolder).ThenBy(e => e.Name).ToList();
    }

    public byte[] Read(string relativePath) => File.ReadAllBytes(Full(relativePath));

    public async Task WriteAsync(string relativePath, Stream content)
    {
        var path = Full(relativePath);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        await using var fs = File.Create(path);
        await content.CopyToAsync(fs);
    }

    public void CreateFolder(string relativeDir, string name)
    {
        if (name.IndexOfAny(['/', '\\']) >= 0) throw new InvalidOperationException("Folder name cannot contain path separators.");
        Directory.CreateDirectory(Path.Combine(Full(relativeDir), name));
    }

    public void Rename(string relativePath, string newName)
    {
        if (newName.IndexOfAny(['/', '\\']) >= 0) throw new InvalidOperationException("Name cannot contain path separators.");
        var source = Full(relativePath);
        var target = Path.Combine(Path.GetDirectoryName(source)!, newName);
        if (Directory.Exists(source)) Directory.Move(source, target);
        else File.Move(source, target);
    }

    public void DeleteFile(string relativePath) => File.Delete(Full(relativePath));

    /// <summary>Deletes a folder only when empty; returns false (with count) otherwise.</summary>
    public bool TryDeleteFolder(string relativePath, out int childCount)
    {
        var path = Full(relativePath);
        childCount = Directory.EnumerateFileSystemEntries(path).Count();
        if (childCount > 0) return false;
        Directory.Delete(path);
        return true;
    }
}
