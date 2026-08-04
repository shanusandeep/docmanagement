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
            settingsPart.Settings = new Settings(new TrackRevisions());
            settingsPart.Settings.Save();
            main.Document.Save();
        }
        return ms.ToArray();
    }
}
