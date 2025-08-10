using System;
using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Linq;

class Program
{
    public static void Main(string[] args)
    {
        string file1 = "1.docx";
        string file2 = "2.docx";
        string output = "merged.docx";
        
        try
        {
            // First, validate the source documents
            Console.WriteLine("Validating source documents...");
            ValidateDocument(file1);
            ValidateDocument(file2);
            
            Console.WriteLine("Creating merged document...");
            CreateSimpleMergedDocument(file1, file2, output);
            Console.WriteLine("Documents merged successfully: " + output);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
            Console.WriteLine("Stack trace: " + ex.StackTrace);
        }
    }

    private static void ValidateDocument(string docPath)
    {
        try
        {
            using (var doc = WordprocessingDocument.Open(docPath, false))
            {
                var mainPart = doc.MainDocumentPart;
                if (mainPart?.Document?.Body == null)
                {
                    throw new InvalidOperationException($"Document {docPath} has invalid structure");
                }
                Console.WriteLine($"✓ {docPath} is valid");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"✗ {docPath} is invalid: {ex.Message}");
            throw;
        }
    }

    private static void CreateSimpleMergedDocument(string firstDoc, string secondDoc, string outputDoc)
    {
        // Delete existing output file
        if (File.Exists(outputDoc))
        {
            File.Delete(outputDoc);
        }

        // Create a new document with minimal structure
        using (var newDoc = WordprocessingDocument.Create(outputDoc, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            // Add main document part with basic structure
            var mainPart = newDoc.AddMainDocumentPart();
            
            // Create document with proper structure
            mainPart.Document = new Document();
            var body = new Body();
            mainPart.Document.Append(body);

            // Add content from first document
            Console.WriteLine("Adding content from first document...");
            AddSimpleDocumentContent(firstDoc, body);

            // Add a simple separator paragraph
            var separator = new Paragraph(
                new Run(
                    new Text("--- Document 2 ---")
                )
            );
            body.Append(separator);

            // Add content from second document
            Console.WriteLine("Adding content from second document...");
            AddSimpleDocumentContent(secondDoc, body);

            // Save with explicit call
            Console.WriteLine("Saving document...");
            mainPart.Document.Save();
        }
    }

    private static void AddSimpleDocumentContent(string sourceDocPath, Body targetBody)
    {
        using (var sourceDoc = WordprocessingDocument.Open(sourceDocPath, false))
        {
            var sourceBody = sourceDoc.MainDocumentPart?.Document?.Body;
            if (sourceBody != null)
            {
                int elementCount = 0;
                foreach (var element in sourceBody.Elements())
                {
                    // Only copy paragraphs - the simplest approach
                    if (element is Paragraph paragraph)
                    {
                        // Create a new simple paragraph with just the text
                        var newParagraph = new Paragraph();
                        var run = new Run();
                        
                        // Extract just the text content, ignoring complex formatting
                        var textContent = paragraph.InnerText;
                        if (!string.IsNullOrWhiteSpace(textContent))
                        {
                            run.Append(new Text(textContent));
                            newParagraph.Append(run);
                            targetBody.Append(newParagraph);
                            elementCount++;
                        }
                    }
                }
                Console.WriteLine($"Added {elementCount} paragraphs from {Path.GetFileName(sourceDocPath)}");
            }
        }
    }
}