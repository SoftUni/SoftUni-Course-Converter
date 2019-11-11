using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using System.Reflection;

public class SoftUniPowerPointConverter
{
    static readonly string pptTemplateFileName = Path.GetFullPath(Directory.GetCurrentDirectory() + 
        @"\..\..\..\Document-Templates\SoftUni-PowerPoint-Template-Nov-2019.pptx");
    static readonly string pptSourceFileName = Directory.GetCurrentDirectory() + @"\test1.pptx";
    static readonly string pptDestFileName = Directory.GetCurrentDirectory() + @"\converted.pptx";

    enum Language { EN, BG };

    static void Main()
    {
        ConvertAndFixPresentation(pptSourceFileName, pptDestFileName, pptTemplateFileName, true);
    }

    public static void ConvertAndFixPresentation(string pptSourceFileName, 
        string pptDestFileName, string pptTemplateFileName, bool appWindowVisible)
    {
        Microsoft.Office.Core.MsoTriState pptAppWindowsVisible = appWindowVisible ?
            Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
        Application pptApp = new Application();
        try
        {
            Console.WriteLine("Processing input presentation: {0}", pptSourceFileName);
            Presentation pptSource = pptApp.Presentations.Open(
                pptSourceFileName, WithWindow: pptAppWindowsVisible);

            Console.WriteLine("Copying the PPTX template '{0}' as output presentation '{1}'", 
                pptTemplateFileName, pptDestFileName);
            File.Copy(pptTemplateFileName, pptDestFileName, true);

            Console.WriteLine($"Opening the output presentation: {pptDestFileName}...");
            Presentation pptDestination = pptApp.Presentations.Open(
                pptDestFileName, WithWindow: pptAppWindowsVisible);

            List<string> pptTemplateSlideTitles = ExtractSlideTitles(pptDestination);

            RemoveAllSectionsAndSlides(pptDestination);

            CopyDocumentProperties(pptSource, pptDestination);

            CopySlidesAndSections(pptSource, pptDestination);
            
            pptSource.Close();

            Language lang = DetectPresentationLanguage(pptDestination);

            FixQuestionsSlide(pptDestination, pptTemplateFileName, pptTemplateSlideTitles, lang);

            FixLicenseSlide(pptDestination, pptTemplateFileName, pptTemplateSlideTitles, lang);

            FixAboutSoftUniSlide(pptDestination, pptTemplateFileName, pptTemplateSlideTitles, lang);

            FixInvalidSlideLayouts(pptDestination);

            FixSectionTitleSlides(pptDestination);

            ReplaceIncorrectHyperlinks(pptDestination);

            FixSlideTitles(pptDestination);

            FixSlideNumbers(pptDestination);

            FixSlideNotesPages(pptDestination);

            pptDestination.Save();
        }
        finally
        {
            if (!appWindowVisible)
                pptApp.Quit();
        }
    }

    static List<Shape> ExtractSlideTitleShapes(Presentation presentation, bool includeSubtitles = false)
    {
        List<Shape> slideTitleShapes = new List<Shape>();
        foreach (Slide slide in presentation.Slides)
        {
            Shape slideTitleShape = null;
            foreach (Shape shape in slide.Shapes.Placeholders)
            {
                if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoPlaceholder)
                    if (shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderTitle)
                        if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                            slideTitleShape = shape;
            }
            if (slideTitleShape == null)
            {
                if (slide.Shapes.Placeholders.Count > 0)
                {
                    Shape firstShape = slide.Shapes.Placeholders[1];
                    if (firstShape?.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                        slideTitleShape = firstShape;
                }
            }
            slideTitleShapes.Add(slideTitleShape);

            if (includeSubtitles)
            {
                // Extract also subtitles
                foreach (Shape shape in slide.Shapes.Placeholders)
                {
                    if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoPlaceholder)
                        if (shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderSubtitle)
                            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                                slideTitleShapes.Add(shape);
                }
            }
        }
        return slideTitleShapes;
    }

    static List<string> ExtractSlideTitles(Presentation presentation)
    {
        List<Shape> slideTitleShapes = ExtractSlideTitleShapes(presentation);
        List<string> slideTitles = slideTitleShapes
            .Select(shape => shape?.TextFrame.TextRange.Text)
            .ToList();
        return slideTitles;
    }

    static void RemoveAllSectionsAndSlides(Presentation presentation)
    {
        Console.WriteLine("Removing all sections and slides from the template...");
        while (presentation.SectionProperties.Count > 0)
            presentation.SectionProperties.Delete(1, true);
    }

    static void CopySlidesAndSections(Presentation pptSource, Presentation pptDestination)
    {
        Console.WriteLine("Copying all slides and sections from the source presentation...");

        // Copy all slides from the source presentation
        Console.WriteLine("  Copying all slides from the source presentation...");
        pptSource.Slides.Range().Copy();
        pptDestination.Slides.Paste();

        // Fix broken source code boxes
        Console.WriteLine("  Fixing source code boxes...");

        int slidesCount = pptSource.Slides.Count;
        for (int slideNum = 1; slideNum <= slidesCount; slideNum++)
        {
            Slide newSlide = pptDestination.Slides[slideNum];
            if (newSlide.CustomLayout.Name == "Source Code Example")
            {
                Slide oldSlide = pptSource.Slides[slideNum];
                for (int shapeNum = 1; shapeNum <= newSlide.Shapes.Placeholders.Count; shapeNum++)
                {
                    Shape newShape = newSlide.Shapes.Placeholders[shapeNum];
                    if (newShape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue &&
                        newShape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        // Found [Code Box] -> copy the paragraph formatting from the original shape
                        Shape oldShape = oldSlide.Shapes.Placeholders[shapeNum];
                        newShape.TextFrame.TextRange.ParagraphFormat.SpaceBefore =
                            Math.Max(0, oldShape.TextFrame.TextRange.ParagraphFormat.SpaceBefore);
                        newShape.TextFrame.TextRange.ParagraphFormat.SpaceAfter =
                            Math.Max(0, oldShape.TextFrame.TextRange.ParagraphFormat.SpaceAfter);
                        newShape.TextFrame.TextRange.ParagraphFormat.SpaceWithin =
                            Math.Max(0, oldShape.TextFrame.TextRange.ParagraphFormat.SpaceWithin);
                        newShape.TextFrame.TextRange.LanguageID =
                            Microsoft.Office.Core.MsoLanguageID.msoLanguageIDEnglishUS;
                        // newShape.TextFrame.TextRange.NoProofing =
                        //    Microsoft.Office.Core.MsoTriState.msoTrue;
                    }
                }

                Console.WriteLine($"    Fixed the code box styling at slide #{slideNum}");
            }
        }

        // Copy all sections from the source presentation
        Console.WriteLine("  Copying all sections from the source presentation...");
        for (int sectNum = 1; sectNum <= pptSource.SectionProperties.Count; sectNum++)
        {
            string sectionName = pptSource.SectionProperties.Name(sectNum);
            sectionName = FixEnglishTitle(sectionName);
            int firstSlide = pptSource.SectionProperties.FirstSlide(sectNum);
            pptDestination.SectionProperties.AddBeforeSlide(firstSlide, sectionName);
        }
    }

    static void CopyDocumentProperties(Presentation pptSource, Presentation pptDestination)
    {
        Console.WriteLine("Copying document properties (metadata)...");

        object srcDocProperties = pptSource.BuiltInDocumentProperties;
        string title = GetObjectProperty(srcDocProperties, "Title")?.ToString();
        string subject = GetObjectProperty(srcDocProperties, "Subject")?.ToString();
        string category = GetObjectProperty(srcDocProperties, "Category")?.ToString();
        string keywords = GetObjectProperty(srcDocProperties, "Keywords")?.ToString();

        object destDocProperties = pptDestination.BuiltInDocumentProperties;
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

    static void SetObjectProperty(object obj, string propName, object propValue)
    {
        object prop = obj.GetType().InvokeMember(
            "Item", BindingFlags.Default | BindingFlags.GetProperty,
            null, obj, new object[] { propName });
        prop.GetType().InvokeMember(
            "Value", BindingFlags.Default | BindingFlags.SetProperty,
            null, prop, new object[] { propValue });
    }

    static void FixInvalidSlideLayouts(Presentation presentation)
    {
        Console.WriteLine("Fixing the invalid slide layouts...");

        var layoutMappings = new Dictionary<string, string> {
            { "Presentation Title Slide", "Presentation Title Slide" },
            { "Section Title Slide", "Section Title Slide" },
            { "Important Concept", "Important Concept" },
            { "Important Example", "Important Example" },
            { "Table of Content", "Table of Contents" },
            { "Comparison Slide", "Comparison Slide" },
            { "Title and Content", "Title and Content" },
            { "Source Code Example", "Source Code Example" },
            { "Image and Content", "Image and Content" },
            { "Questions Slide", "Questions Slide" },
            { "Blank Slide", "Blank Slide" },
            { "About Slide", "About Slide" },
            { "Заглавие и съдържание", "Title and Content" },
            { "Section Slide", "Section Title Slide" },
            { "Title Slide", "Section Title Slide" },
            { "Заглавен слайд", "Section Title Slide" },
            { "Block scheme", "Blank Slide" },
            { "Comparison Slide Dark", "Comparison Slide" },
            { "Last", "About Slide" },
            { "", "" },
        };
        const string defaultLayoutName = "Title and Content";

        var customLayoutsByName = new Dictionary<string, CustomLayout>();
        foreach (CustomLayout layout in presentation.SlideMaster.CustomLayouts)
            customLayoutsByName[layout.Name] = layout;
        var layoutsForDeleting = new HashSet<string>();

        // Replace the incorrect layouts with the correct ones
        for (int slideNum = 1; slideNum <= presentation.Slides.Count; slideNum++)
        {
            Slide slide = presentation.Slides[slideNum];
            string oldLayoutName = slide.CustomLayout.Name;
            string newLayoutName = defaultLayoutName;
            if (layoutMappings.ContainsKey(oldLayoutName))
                newLayoutName = layoutMappings[oldLayoutName];
            if (newLayoutName != oldLayoutName)
            {
                Console.WriteLine($"  Replacing invalid slide layout \"{oldLayoutName}\" on slide #{slideNum} with \"{newLayoutName}\"");
                // Replace the old layout with the new layout
                slide.CustomLayout = customLayoutsByName[newLayoutName];
                layoutsForDeleting.Add(oldLayoutName);
            }
        }

        // Delete all old (and no longer used) layouts
        foreach (var layoutName in layoutsForDeleting)
        {
            Console.WriteLine($"  Deleting unused layout \"{layoutName}\"");
            CustomLayout layout = customLayoutsByName[layoutName];
            layout.Delete();
        }
    }

    static void ReplaceIncorrectHyperlinks(Presentation presentation)
    {
        var hyperlinksToReplace = new Dictionary<string, string> {
            { "http://softuni.bg", "https://softuni.bg" },
            { "http://softuni.foundation/", "https://softuni.foundation" },
        };

        Console.WriteLine("Replacing incorrect hyperlinks in the slides...");
        for (int slideNum = 1; slideNum <= presentation.Slides.Count; slideNum++)
        {
            Slide slide = presentation.Slides[slideNum];
            foreach (Hyperlink link in slide.Hyperlinks)
            {
                try
                {
                    string linkText = link.TextToDisplay;
                    if (hyperlinksToReplace.ContainsKey(linkText))
                    {
                        string newText = hyperlinksToReplace[linkText];
                        link.TextToDisplay = newText;
                        link.Address = newText;
                    }
                }
                catch (Exception)
                {
                    // Ignore silently: cannot change the link for some reason
                }
            }
        }
    }

    static void FixSectionTitleSlides(Presentation presentation)
    {
        Console.WriteLine("Fixing broken section title slides...");

        var sectionTitleSlides = presentation.Slides.Cast<Slide>()
            .Where(slide => slide.CustomLayout.Name == "Section Title Slide");
        foreach (Slide slide in sectionTitleSlides)
        {
            // Collect the texts from the slide (expecting title and subtitle)
            List<string> slideTexts = new List<string>();
            foreach (Shape shape in slide.Shapes)
            {
                try
                {
                    if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue
                        && (shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderTitle
                            || shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderSubtitle
                            || shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderBody)
                        && shape.TextFrame.TextRange.Text != "")
                    {
                        slideTexts.Add(shape.TextFrame.TextRange.Text);
                        shape.Delete();
                    }
                }
                catch (Exception)
                {
                    // Silently ignore --> the shape is not a placeholder
                }
            }

            // Put the slide texts into the placeholders (and delete the empty placeholders)
            for (int i = 0; i < slide.Shapes.Placeholders.Count; i++)
            {
                Shape placeholder = slide.Shapes.Placeholders[i+1];
                if (i < slideTexts.Count)
                    placeholder.TextFrame.TextRange.Text = slideTexts[i];
                else
                    placeholder.Delete();
            }
            Console.WriteLine($" Fixed slide #{slide.SlideNumber}: {slideTexts.FirstOrDefault()}");
        }
    }

    static void FixSlideTitles(Presentation presentation)
    {
        Console.WriteLine("Fixing incorrect slide titles...");

        var titleMappings = new Dictionary<string, string> {
            { "Table of Content", "Table of Contents" }
        };

        List<Shape> slideTitleShapes = 
            ExtractSlideTitleShapes(presentation, includeSubtitles: true);
        List<string> slideTitles = slideTitleShapes
            .Select(shape => shape?.TextFrame.TextRange.Text)
            .ToList();
        for (int i = 0; i < slideTitleShapes.Count; i++)
        {
            string newTitle = FixEnglishTitle(slideTitles[i]);
            if (titleMappings.ContainsKey(newTitle))
                newTitle = titleMappings[newTitle];
            if (newTitle != slideTitles[i])
            {
                Console.WriteLine($"  Replaced slide #{i} title: [{slideTitles[i]}] -> [{newTitle}]");
                slideTitleShapes[i].TextFrame.TextRange.Text = newTitle;
            }
        }
    }

    static readonly HashSet<string> EnglishTitleCaseIgnoredWords = new HashSet<string> {
        "a", "an", "the", "is", "vs",
        "and", "or", "in", "of", "by", "from", "at", "off", "to",
        "into", "about", "onto", "for", "with"
    };

    static string FixEnglishTitle(string text)
    {
        string EnglishWordToTitleCase(string word)
        {
            if (string.IsNullOrEmpty(word))
                return word;

            // Handle normal words like "program"
            if (char.ToLower(word[0]) >= 'a' && char.ToLower(word[0]) <= 'z')
                return "" + char.ToUpper(word[0]) + word.Substring(1);

            // Handle words like "[Run]" or "(maybe)"
            if (word.Length > 1 && char.ToLower(word[1]) >= 'a' && char.ToLower(word[1]) <= 'z')
                return "" + word[0] + char.ToUpper(word[1]) + word.Substring(2);

            return word;
        }

        if (string.IsNullOrEmpty(text))
            return text;

        text = text.Replace(" - ", " – ");

        string[] words = text.Split(' ');
        if (words.Length > 0)
        {
            // Always start with capital letter
            words[0] = EnglishWordToTitleCase(words[0]);
        }
        for (int i = 1; i < words.Length; i++)
        {
            string wordOnly = words[i].Trim(' ', ',', ';', '?', '!', '.', '(', ')');
            if (EnglishTitleCaseIgnoredWords.Contains(wordOnly.ToLower()))
            {
                // Special word (like preposition / conjunctions) -> lowercase it (unless it is ALL CAPS)
                if (wordOnly != wordOnly.ToUpper())
                    words[i] = words[i].ToLower();
            }
            else
            {
                // Normal word (non-special) -> capitalize its first letter
                words[i] = EnglishWordToTitleCase(words[i]);
            }
        }

        string result = string.Join(" ", words);
        return result;
    }

    static Language DetectPresentationLanguage(Presentation presentation)
    {
        var englishLettersCount = 0;
        var bulgarianLettersCount = 0;
        var slideTitles = ExtractSlideTitles(presentation);
        foreach (string title in slideTitles)
            if (title != null)
                foreach (char ch in title.ToLower())
                    if (ch >= 'a' && ch <= 'z')
                        englishLettersCount++;
                    else if (ch >= 'а' && ch <= 'я')
                        bulgarianLettersCount++;

        if (bulgarianLettersCount > englishLettersCount / 2)
            return Language.BG;
        else
            return Language.EN;
    }

    static void FixQuestionsSlide(Presentation presentation, string pptTemplateFileName,
        List<string> pptTemplateSlideTitles, Language lang)
    {
        Console.WriteLine("Fixing the [Questions] slide...");
        string questionsSlideTitle =
            (lang == Language.EN) ? "Questions?" : "Въпроси?";
        for (int slideNum = 1; slideNum <= presentation.Slides.Count; slideNum++)
        {
            Slide slide = presentation.Slides[slideNum];
            if (slide.CustomLayout.Name == "1_Questions Slide")
            {
                Console.WriteLine($"  Found the [Questions] slide #{slideNum} --> replacing it from the template");

                presentation.Slides[slideNum].Delete();
                int questionsSlideIndexInTemplate = pptTemplateSlideTitles.LastIndexOf(questionsSlideTitle);
                presentation.Slides.InsertFromFile(pptTemplateFileName, slideNum - 1,
                    questionsSlideIndexInTemplate + 1, questionsSlideIndexInTemplate + 1);
            }
        }
    }

    static void FixLicenseSlide(Presentation presentation, string pptTemplateFileName,
        List<string> pptTemplateSlideTitles, Language lang)
    {
        Console.WriteLine("Fixing the [License] slide...");
        string licenseSlideTitle =
            (lang == Language.EN) ? "License" : "Лиценз";
        var slideTitles = ExtractSlideTitles(presentation);
        for (int slideNum = 1; slideNum <= presentation.Slides.Count; slideNum++)
        {
            if (slideTitles[slideNum - 1] == "License" ||
                slideTitles[slideNum - 1] == "Лиценз")
            {
                Console.WriteLine($"  Found the [License] slide #{slideNum} --> replacing it from the template");

                presentation.Slides[slideNum].Delete();
                int licenseSlideIndexInTemplate = pptTemplateSlideTitles.LastIndexOf(licenseSlideTitle);
                presentation.Slides.InsertFromFile(pptTemplateFileName, slideNum - 1,
                    licenseSlideIndexInTemplate + 1, licenseSlideIndexInTemplate + 1);
            }
        }
    }

    static void FixAboutSoftUniSlide(Presentation presentation, string pptTemplateFileName,
        List<string> pptTemplateSlideTitles, Language lang)
    {
        Console.WriteLine("Fixing the [About] slide...");

        List<string> aboutSlidePossibleTitles;
        if (lang == Language.EN)
            aboutSlidePossibleTitles = new List<string>() 
            {
                "Trainings @ Software University (SoftUni)",
                "About SoftUni Digital",
                "About SoftUni Creative",
                "About SoftUni Kids"
            };
        else if (lang == Language.BG)
            aboutSlidePossibleTitles = new List<string>() 
            {
                "Обучения в СофтУни",
                "Обучения в Софтуерен университет (СофтУни)"
            };
        else
            throw new ArgumentException("Invalid language");
        int aboutSlideIndexInTemplate = -1;
        foreach (var title in aboutSlidePossibleTitles)
            aboutSlideIndexInTemplate = Math.Max(aboutSlideIndexInTemplate,
                pptTemplateSlideTitles.LastIndexOf(title));

        var slideTitles = ExtractSlideTitles(presentation);
        for (int slideNum = 1; slideNum <= presentation.Slides.Count; slideNum++)
        {
            if (aboutSlidePossibleTitles.Contains(slideTitles[slideNum - 1]))
            {
                Console.WriteLine($"  Found the [About] slide #{slideNum} --> replacing it from the template");
                presentation.Slides[slideNum].Delete();
                presentation.Slides.InsertFromFile(pptTemplateFileName, slideNum - 1,
                    aboutSlideIndexInTemplate + 1, aboutSlideIndexInTemplate + 1);
            }
        }
    }

    static void FixSlideNumbers(Presentation presentation)
    {
        Shape FindFirstSlideNumberShape()
        {
            CustomLayout layout =
                presentation.SlideMaster.CustomLayouts.OfType<CustomLayout>()
                .Where(l => l.Name == "Title and Content").First();
            foreach (Shape shape in layout.Shapes)
                if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoPlaceholder)
                    if (shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderSlideNumber)
                        return shape;
            return null;
        }

        Console.WriteLine("Fixing the slide numbering...");

        var layoutsWithoutNumbering = new HashSet<string>() {
            "Presentation Title Slide",
            "Section Title Slide",
            "Questions Slide"
        };

        Shape slideNumberShape = FindFirstSlideNumberShape();
        slideNumberShape.Copy();

        // Delete the [slide number] box in each slide, then put it again if needed
        for (int slideNum = 1; slideNum <= presentation.Slides.Count; slideNum++)
        {
            Slide slide = presentation.Slides[slideNum];
            string layoutName = slide.CustomLayout.Name;

            foreach (Shape shape in slide.Shapes)
            {
                bool isSlideNumberTextBox =
                    shape.Type == Microsoft.Office.Core.MsoShapeType.msoTextBox
                    && shape.Name.Contains("Slide Number");
                bool isSlideNumberPlaceholder =
                    shape.Type == Microsoft.Office.Core.MsoShapeType.msoPlaceholder
                    && shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderSlideNumber;
                if (isSlideNumberTextBox || isSlideNumberPlaceholder)
                {
                    // Found a "slide number" shape --> delete it
                    shape.Delete();
                }
            }

            if (!layoutsWithoutNumbering.Contains(layoutName))
            {
                // The slide should have [slide number] box --> insert it
                slide.Shapes.Paste();
            }
        }
    }

    static void FixSlideNotesPages(Presentation presentation)
    {
        Shape FindNotesFooter()
        {
            var footerShape = presentation.NotesMaster.Shapes.OfType<Shape>().Where(
                shape => shape.Type == Microsoft.Office.Core.MsoShapeType.msoPlaceholder
                && shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderFooter)
                .FirstOrDefault();
            return footerShape;
        }

        Shape footerFromNotesMaster = FindNotesFooter();
        footerFromNotesMaster.Copy();

        for (int slideNum = 1; slideNum <= presentation.Slides.Count; slideNum++)
        {
            Slide slide = presentation.Slides[slideNum];
            if (slide.HasNotesPage == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                var slideNotesFooter = slide.NotesPage.Shapes.OfType<Shape>().Where(
                    shape => shape.Type == Microsoft.Office.Core.MsoShapeType.msoPlaceholder
                    && shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderFooter)
                    .FirstOrDefault();
                if (slideNotesFooter != null)
                {
                    slideNotesFooter.Delete();
                }
                slide.NotesPage.Shapes.Paste();
            }
        }
    }
}
