using System;
using System.IO;
using Microsoft.Office.Interop.Word;

public class SoftUniMSWordConverter
{
    const string docTemplateFileName = @"C:\Users\nakov\Desktop\SoftUni-Course-Catalog\Document-Templates\SoftUni-Creative-Document-Template-Nov-2019.docx";
    static readonly string docSourceFileName = Directory.GetCurrentDirectory() + @"\test1.docx";
    static readonly string docDestFileName = Directory.GetCurrentDirectory() + @"\converted.docx";

    static void Main()
    {
        ConvertAndFixDocument(docSourceFileName, docDestFileName, docTemplateFileName, false);
    }

    public static void ConvertAndFixDocument(string docSourceFileName,
        string docDestFileName, string docTemplateFileName, bool appWindowVisible)
    {
        Application wordApp = new Application();
        wordApp.Visible = appWindowVisible; // Show / hide MS Word app window
        wordApp.ScreenUpdating = appWindowVisible; // Enable / disable screen updates after each change

        try
        {
            Console.WriteLine("Processing: {0}", docSourceFileName);
            Document docSource = wordApp.Documents.Open(docSourceFileName);

            Console.WriteLine("Loading template: {0}", docTemplateFileName);
            Document docTemplate = wordApp.Documents.Open(docTemplateFileName);

            Document docDestination = docTemplate;

            CopyDocumentContent(docSource, docDestination);

            CopyDocumentProperties(docSource, docDestination);

            docDestination.SaveAs(docDestFileName);
            docSource.Close();
        }
        finally
        {
            if (!appWindowVisible)
                wordApp.Quit(false);
        }
    }

    static void CopyDocumentProperties(Document docSource, Document docDestination)
    {
    }

    static void CopyDocumentContent(Document docSource, Document docDestination)
    {
        Console.WriteLine("Copying the entire content from the source document...");
        docSource.Content.Copy();
        docDestination.Content.Paste();
    }
}
