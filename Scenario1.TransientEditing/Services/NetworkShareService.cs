using Microsoft.Extensions.Options;

namespace Scenario1.TransientEditing.Services;

/// <summary>
/// Simulates the corporate network share (Serfis document storage).
/// Points at a real UNC path in production; a local folder for the demo.
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
        if (Directory.EnumerateFiles(_root, "*.docx").Any()) return;
        foreach (var name in new[]
                 {
                     "Case-1001 Investigation Report",
                     "Case-1002 Interview Notes",
                     "Case-1003 Evidence Summary"
                 })
        {
            File.WriteAllBytes(Path.Combine(_root, name + ".docx"),
                _word.CreateBlankDocx($"{name} (created by Serfis on the network share)"));
        }
    }

    public List<ShareFile> List() =>
        Directory.EnumerateFiles(_root, "*.docx")
            .Select(f => new FileInfo(f))
            .Select(fi => new ShareFile
            {
                Name = fi.Name,
                SizeBytes = fi.Length,
                ModifiedUtc = fi.LastWriteTimeUtc
            })
            .OrderBy(f => f.Name)
            .ToList();

    public byte[] Read(string fileName) => File.ReadAllBytes(Path.Combine(_root, fileName));

    public async Task WriteAsync(string fileName, Stream content)
    {
        await using var fs = File.Create(Path.Combine(_root, fileName));
        await content.CopyToAsync(fs);
    }

    public void CreateNew(string fileName)
    {
        if (!fileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)) fileName += ".docx";
        File.WriteAllBytes(Path.Combine(_root, fileName),
            _word.CreateBlankDocx(Path.GetFileNameWithoutExtension(fileName)));
    }
}
