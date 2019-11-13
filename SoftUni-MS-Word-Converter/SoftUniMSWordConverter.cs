using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using static SoftUniConverterCommon.ConverterUtils;

public class SoftUniMSWordConverter
{
    static readonly string docTemplateFileName = Path.GetFullPath(Directory.GetCurrentDirectory() + 
        @"\..\..\..\Document-Templates\SoftUni-Creative-Document-Template-Nov-2019.docx");
    static readonly string docSourceFileName = Path.GetFullPath(Directory.GetCurrentDirectory() +
        @"\..\..\..\Sample-Docs\test3.docx");
    static readonly string docDestFileName = Directory.GetCurrentDirectory() + @"\converted.docx";

    static void Main()
    {
        ConvertAndFixDocument(docSourceFileName, docDestFileName, docTemplateFileName, false);
    }

    public static void ConvertAndFixDocument(string docSourceFileName,
        string docDestFileName, string docTemplateFileName, bool appWindowVisible)
    {
        if (KillAllProcesses("WINWORD"))
            Console.WriteLine("MS Word (WINWORD.EXE) is still running -> process terminated.");

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
            
            docSource.Close(false);

            FixDocumentStylesAndHeadings(docDestination);

            FixWordsLanguage(docDestination);

            docDestination.Save();

            if (!appWindowVisible)
                docDestination.Close();
        }
        finally
        {
            if (!appWindowVisible)
            {
                // Quit the MS Word application
                wordApp.Quit(false);

                // Release any associated .NET proxies for the COM objects, which are not in use
                // Intentionally we call the garbace collector twice
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
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

    static void FixDocumentStylesAndHeadings(Document doc)
    {
        Console.WriteLine("Fixing incorrect document styles and headings...");

        var styleMappings = new Dictionary<string, string> {
            { "Heading", "Heading 1" },
        };

        var headingStyleNames = new HashSet<string>() { "Heading 1", "Heading 2",
            "Heading 3", "Heading 4", "Heading 5", "Heading 6"};
        foreach (var oldStyle in styleMappings.Keys)
            headingStyleNames.Add(oldStyle);

        // Replace the incorrect styles with the correct ones
        Dictionary<string, Style> stylesByName = null;
        foreach (Paragraph paragraph in doc.Paragraphs)
        {
            // Replace incorrect styles by corresponding correct style
            Style paragraphStyle = paragraph.get_Style();
            string oldStyleName = paragraphStyle.NameLocal;
            if (styleMappings.ContainsKey(oldStyleName))
            {
                // Replace the old (incorrect) style with the new (correct) style
                if (stylesByName == null)
                {
                    Console.WriteLine("  Reading document styles...");
                    stylesByName = new Dictionary<string, Style>();
                    foreach (Style docStyle in doc.Styles)
                        stylesByName[docStyle.NameLocal] = docStyle;
                }
                string newStyleName = styleMappings[oldStyleName];
                Style newStyle = stylesByName[newStyleName];
                string paragraphText = TruncateString(paragraph.Range.Text, 50);
                Console.WriteLine($"  Replacing invalid style '{oldStyleName}' with correct style '{newStyleName}' in paragraph '{paragraphText}'");
                paragraph.set_Style(newStyle);
            }

            // Fix document headings to use "Title Case"
            if (headingStyleNames.Contains(oldStyleName))
            {
                string oldHeadingText = paragraph.Range.Text;
                string newHeadingText = FixEnglishTitleCharacterCasing(oldHeadingText);
                if (newHeadingText != oldHeadingText)
                {
                    paragraph.Range.Text = newHeadingText;
                    Console.WriteLine($"  Replacing heading: '{oldHeadingText.Trim()}' -> '{newHeadingText.Trim()}'");
                }
            }

            Console.Write("."); // Display progress of the current operation
        }

        Console.WriteLine();

        // Delete all old (and no longer used) styles
        foreach (var styleName in styleMappings.Keys)
        {
            if (stylesByName != null && stylesByName.ContainsKey(styleName))
            {
                Console.WriteLine($"  Deleting unused style '{styleName}'...");
                Style style = stylesByName[styleName];
                style.Delete();
            }
        }
    }

    static void FixWordsLanguage(Document doc)
    {
        Console.WriteLine("Fixing words language...");

        foreach (Range word in doc.Words)
        {
            Console.Write("."); // Show operation progress
            string wordText = word.Text.Trim();
            var cyrillicLettersCount = CountCyrillicLetters(wordText);
            if (cyrillicLettersCount == wordText.Length)
            {
                // The word holds Cyrillic letters only --> set Bulgarian language
                word.LanguageID = WdLanguageID.wdBulgarian;
            }
            else
            {
                // The word holds non-Cyrillic letters --> set English language
                word.LanguageID = WdLanguageID.wdEnglishUS;
                bool isCodeIdentifier = IsCodeIdentifier(wordText);
                bool isBracket = (wordText == "(") || (wordText == ")");
                bool isMonospacedFont = (word.Font.Name == "Consolas");
                if (isCodeIdentifier || isBracket || isMonospacedFont)
                {
                    // Disable the spell checker for code identifiers
                    word.NoProofing = (int)MsoTriState.msoTrue;
                }
            }
        }

        int CountCyrillicLetters(string word) =>
            word.ToLower().Count(l => l >= 'а' && l <= 'я');

        bool IsCodeIdentifier(string word) =>
            word.Skip(1).Count(l => l >= 'A' && l <= 'Z') > 0;

        Console.WriteLine();
    }
}
