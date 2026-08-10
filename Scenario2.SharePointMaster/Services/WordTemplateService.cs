using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Scenario2.SharePointMaster.Services;

/// <summary>
/// Creates blank .docx files with Track Changes enforced by injecting
/// &lt;w:trackChanges/&gt; (SDK class TrackRevisions) into word/settings.xml.
/// </summary>
public class WordTemplateService
{
    /// <summary>Returns the same document with Track Changes forced on.</summary>
    public byte[] EnsureTrackChanges(byte[] docxBytes)
    {
        using var ms = new MemoryStream();
        ms.Write(docxBytes, 0, docxBytes.Length);
        ms.Position = 0;
        using (var doc = WordprocessingDocument.Open(ms, true))
        {
            if (doc.MainDocumentPart is { } main)
            {
                var settingsPart = main.DocumentSettingsPart ?? main.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings ??= new Settings();
                if (!settingsPart.Settings.Elements<TrackRevisions>().Any())
                    settingsPart.Settings.PrependChild(new TrackRevisions());
                if (!settingsPart.Settings.Elements<Compatibility>().Any())
                {
                    // declare modern Word format so documents don't open in Compatibility Mode
                    settingsPart.Settings.AppendChild(new Compatibility(new CompatibilitySetting
                    {
                        Name = CompatSettingNameValues.CompatibilityMode,
                        Uri = "http://schemas.microsoft.com/office/word",
                        Val = "15"
                    }));
                }
                settingsPart.Settings.Save();
            }
        }
        return ms.ToArray();
    }

    public byte[] CreateBlankDocx(string title)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text(title))),
                new Paragraph(new Run(new Text("")))));
            var settingsPart = main.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Settings(
                new TrackRevisions(),
                new Compatibility(new CompatibilitySetting
                {
                    Name = CompatSettingNameValues.CompatibilityMode,
                    Uri = "http://schemas.microsoft.com/office/word",
                    Val = "15"
                }));
            settingsPart.Settings.Save();
            main.Document.Save();
        }
        return ms.ToArray();
    }
}
