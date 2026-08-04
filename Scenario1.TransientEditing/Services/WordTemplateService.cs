using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Scenario1.TransientEditing.Services;

/// <summary>
/// Creates blank .docx files and enforces Track Changes by injecting
/// &lt;w:trackChanges/&gt; into word/settings.xml (see design doc, Section 6).
/// </summary>
public class WordTemplateService
{
    public byte[] CreateBlankDocx(string title)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text(title))),
                new Paragraph(new Run(new Text("")))));
            EnsureTrackChangesPart(main);
            main.Document.Save();
        }
        return ms.ToArray();
    }

    /// <summary>Returns the same document with Track Changes forced on.</summary>
    public byte[] EnsureTrackChanges(byte[] docxBytes)
    {
        using var ms = new MemoryStream();
        ms.Write(docxBytes, 0, docxBytes.Length);
        ms.Position = 0;
        using (var doc = WordprocessingDocument.Open(ms, true))
        {
            if (doc.MainDocumentPart is { } main) EnsureTrackChangesPart(main);
        }
        return ms.ToArray();
    }

    private static void EnsureTrackChangesPart(MainDocumentPart main)
    {
        var settingsPart = main.DocumentSettingsPart ?? main.AddNewPart<DocumentSettingsPart>();
        settingsPart.Settings ??= new Settings();
        // TrackRevisions is the SDK class for the <w:trackChanges/> settings element
        if (!settingsPart.Settings.Elements<TrackRevisions>().Any())
            settingsPart.Settings.PrependChild(new TrackRevisions());
        settingsPart.Settings.Save();
    }
}
