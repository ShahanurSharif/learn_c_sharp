using System.IO;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using System.Collections.Generic;
using System.Xml.Linq;

class Program
{

    public static void Main(string[] args)
    {
        string file1 = "1.docx";
        string file2 = "2.docx";
        string output = "merged.docx";

        // Open source files as OpenXmlPackage
        using (var doc1 = WordprocessingDocument.Open(file1, false))
        using (var doc2 = WordprocessingDocument.Open(file2, false))
        {
            // Get the content as XDocument (LINQ to XML)
            var sources = new List<Source>
            {
                new Source(new WmlDocument(file1), true),
                new Source(new WmlDocument(file2), true)
            };

            // Merge using PowerTools
            WmlDocument mergedDoc = DocumentBuilder.BuildDocument(sources);

            // Save the result
            mergedDoc.SaveAs(output);
        }
        System.Console.WriteLine("Merge complete! Saved as: " + output);
        
    }
    
}