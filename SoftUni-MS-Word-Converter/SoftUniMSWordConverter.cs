using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Word;

public class SoftUniMSWordConverter
{
    static readonly string docTemplateFileName = 
        Path.GetFullPath(Directory.GetCurrentDirectory() + 
        @"\..\..\..\Document-Templates\SoftUni-Creative-Document-Template-Nov-2019.docx");
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
            Console.WriteLine("Processing input document: {0}", docSourceFileName);
            Document docSource = wordApp.Documents.Open(docSourceFileName);

            Console.WriteLine("Copying the DOCX template '{0}' as output document '{1}'",
                docTemplateFileName, docDestFileName);
            File.Copy(docTemplateFileName, docDestFileName, true);

            Console.WriteLine($"Opening the output document: {docDestFileName}...");
            Document docDestination = wordApp.Documents.Open(docDestFileName);

            CopyDocumentContent(docSource, docDestination);

            CopyDocumentProperties(docSource, docDestination);

            docSource.Close();

            FixDocumentStyles(docDestination);

            docDestination.Save();
        }
        finally
        {
            if (!appWindowVisible)
                wordApp.Quit(false);
        }
    }

    static void SetObjectProperty(object obj, string propName, object propValue)
    {
        object prop = obj.GetType().InvokeMember(
            "Item", BindingFlags.Default | BindingFlags.GetProperty,
            null, obj, new object[] { propName });
        prop.GetType().InvokeMember(
            "Value", BindingFlags.Default | BindingFlags.SetProperty,
            null, prop, new object[] { propValue });
    }

    static void CopyDocumentContent(Document docSource, Document docDestination)
    {
        Console.WriteLine("Copying the entire content from the source document...");
        docSource.Content.Copy();
        docDestination.Content.Paste();
    }

    static void CopyDocumentProperties(Document docSource, Document docDestination)
    {
        Console.WriteLine("Copying document properties (metadata)...");

        object srcDocProperties = docSource.BuiltInDocumentProperties;
        string title = GetObjectProperty(srcDocProperties, "Title")
            ?.ToString()?.Replace(',', ';');
        string subject = GetObjectProperty(srcDocProperties, "Subject")
            ?.ToString()?.Replace(',', ';');
        string category = GetObjectProperty(srcDocProperties, "Category")
            ?.ToString()?.Replace(',', ';');
        string keywords = GetObjectProperty(srcDocProperties, "Keywords")
            ?.ToString()?.Replace(',', ';');

        object destDocProperties = docDestination.BuiltInDocumentProperties;
        if (!string.IsNullOrWhiteSpace(title))
            SetObjectProperty(destDocProperties, "Title", title);
        if (!string.IsNullOrWhiteSpace(subject))
            SetObjectProperty(destDocProperties, "Subject", subject);
        if (!string.IsNullOrWhiteSpace(category))
            SetObjectProperty(destDocProperties, "Category", category);
        if (!string.IsNullOrWhiteSpace(keywords))
            SetObjectProperty(destDocProperties, "Keywords", keywords);
    }

    static object GetObjectProperty(object obj, string propName)
    {
        object prop = obj.GetType().InvokeMember(
            "Item", BindingFlags.Default | BindingFlags.GetProperty,
            null, obj, new object[] { propName });
        object propValue = prop.GetType().InvokeMember(
            "Value", BindingFlags.Default | BindingFlags.GetProperty,
            null, prop, new object[] { });
        return propValue;
    }

    static void FixDocumentStyles(Document doc)
    {
        Console.WriteLine("Fixing document styles...");

        var styleMappings = new Dictionary<string, string> {
            { "Heading", "Heading 1" },
        };

        var stylesByName = new Dictionary<string, Style>();
        foreach (Style style in doc.Styles)
            stylesByName[style.NameLocal] = style;

        // Replace the incorrect styles with the correct ones
        foreach (Paragraph paragraph in doc.Paragraphs)
        {
            Style style = paragraph.get_Style();
            string oldStyleName = style.NameLocal;
            if (styleMappings.ContainsKey(oldStyleName))
            {
                // Replace the old (incorrect) style with the new (correct) style
                string newStyleName = styleMappings[oldStyleName];
                Style newStyle = stylesByName[newStyleName];
                string paragraphText = TruncateString(paragraph.Range.Text, 50);
                Console.WriteLine($"  Replacing invalid style '{oldStyleName}' with correct style '{newStyleName}' in paragraph '{paragraphText}'");
                paragraph.set_Style(newStyle);
            }
        }

        // Delete all old (and no longer used) styles
        foreach (var styleName in styleMappings.Keys)
        {
            Style style = stylesByName[styleName];
            style.Delete();
        }
    }

    static string TruncateString(string str, int maxLength)
    {
        if (str == null)
            return "";
        str = str.Trim();
        if (str.Length > maxLength)
            str = str.Substring(0, maxLength) + "...";
        return str;
    }
}
