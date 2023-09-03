using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace OpenXml;

public class SdkTests
{
    private const string SrcFile = "../../../../DA001-TemplateDocument.docx";
    
    [Theory]
    [InlineData(MarkupCompatibilityProcessMode.NoProcess)]
    [InlineData(MarkupCompatibilityProcessMode.ProcessAllParts)]
    [InlineData(MarkupCompatibilityProcessMode.ProcessLoadedPartsOnly)]
    public async Task EnsureThatAttrsAreSaved(MarkupCompatibilityProcessMode mode)
    {
        var fileName = $"temp-{mode}.docx";
        await AddAttributesAndSave(mode, fileName);
        ValidateAttributes(fileName);
    }
    
    private async Task AddAttributesAndSave(MarkupCompatibilityProcessMode mode, string dstFile)
    {
        // Load Source File into memory
        using var ms = new MemoryStream();
        var srcStream = File.OpenRead(SrcFile);
        await srcStream.CopyToAsync(ms);
            
        // Add custom attributes to all elements
        ms.Seek(0, SeekOrigin.Begin);
        
        var os = new OpenSettings
        {
            MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(
                mode, FileFormatVersions.Office2007)
        };
        using(var doc = WordprocessingDocument.Open(ms, true, os))
        {
            var mainPart = doc.MainDocumentPart!;
            var xDocument = mainPart.GetXDocument();
            AssignUnidToAllElements(xDocument.Root!);
            IgnorePt14Namespace(xDocument.Root!);
            mainPart.PutXDocument();
        }
        
        // Save to destination file
        ms.Seek(0, SeekOrigin.Begin);
        await using (var fs = File.OpenWrite(dstFile))
        {
            await ms.CopyToAsync(fs);
        }
    }

    static class PtOpenXml
    {
        public static readonly XNamespace Pt = "http://powertools.codeplex.com/2011";
        public static readonly XName Unid = Pt + "Unid";
    }
    static class MC
    {
        static readonly XNamespace Mc =
            "http://schemas.openxmlformats.org/markup-compatibility/2006";
        public static readonly XName Ignorable = Mc + "Ignorable";
    }

    private static void AssignUnidToAllElements(XElement contentElement)
    {
        foreach (var element in contentElement.Descendants())
        {
            if (element.Attribute(PtOpenXml.Unid) is not null) 
                continue;
            
            var unid = Guid.NewGuid().ToString().Replace("-", "");
            var newAtt = new XAttribute(PtOpenXml.Unid, unid);
            element.Add(newAtt);
        }
    }
    
    private static void IgnorePt14Namespace(XElement root)
    {
        if (root.Attribute(XNamespace.Xmlns + "pt14") == null)
        {
            root.Add(new XAttribute(XNamespace.Xmlns + "pt14", PtOpenXml.Pt.NamespaceName));
        }

        var ignorable = (string?) root.Attribute(MC.Ignorable);
        if (ignorable is not null)
        {
            var list = ignorable.Split(' ');
            if (!list.Contains("pt14"))
            {
                ignorable += " pt14";
                root.SetAttributeValue(MC.Ignorable, ignorable);
            }
        }
        else
        {
            root.Add(new XAttribute(MC.Ignorable, "pt14"));
        }
    }

    private void ValidateAttributes(string fileName)
    {
        using var doc = WordprocessingDocument.Open(fileName, false);
        var mainPart = doc.MainDocumentPart!;
        var xDocument = mainPart.GetXDocument();
        
        foreach (var element in xDocument.Root!.Descendants())
        {
            var attr = element.Attribute(PtOpenXml.Unid);
            Assert.NotNull(attr);
        }
    }
}

public static class Helpers {
    public static XDocument GetXDocument(this OpenXmlPart part)
    {
        var partXDocument = part.Annotation<XDocument>();
        if (partXDocument is not null) 
            return partXDocument;

        using (var partStream = part.GetStream())
        {
            using var partXmlReader = XmlReader.Create(partStream);
            partXDocument = XDocument.Load(partXmlReader);
        }

        part.AddAnnotation(partXDocument);
        return partXDocument;
    }
    
    public static void PutXDocument(this OpenXmlPart part)
    {
        var partXDocument = part.GetXDocument();
        
        using var partStream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var partXmlWriter = XmlWriter.Create(partStream);
        partXDocument.Save(partXmlWriter);
    }
}
