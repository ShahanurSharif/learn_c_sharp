using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

class Program
{

    public static void Main(string[] args)
    {
        Console.Write("Enter the first file name: ");
        object fileName1 = Console.ReadLine();
        
        Console.Write("Enter the second file name: ");
        object fileName2 = Console.ReadLine();

        using var doc1 = WordprocessingDocument.Open(fileName1, false);
        var body1 = doc1.MainDocumentPart.Document.Body;
        
    }
    
}