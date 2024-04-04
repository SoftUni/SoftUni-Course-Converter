using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using SoftUniConverterCommon;
using static SoftUniConverterCommon.ConverterUtils;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using System.Security.Policy;

public class SoftUniPowerPointConverter
{
    static readonly string pptTemplateFileName = Path.GetFullPath(Directory.GetCurrentDirectory() + 
        @"\..\..\..\Document-Templates\SoftUni-Creative-PowerPoint-Template-Nov-2019.pptx");
    static readonly string pptSourceFileName = Path.GetFullPath(Directory.GetCurrentDirectory() +
        @"\..\..\..\Sample-Docs\test3.pptx");
    static readonly string pptDestFileName = Directory.GetCurrentDirectory() + @"\converted.pptx";

    static void Main()
    {
        ConvertAndFixPresentation(pptSourceFileName, pptDestFileName, pptTemplateFileName, true);
    }

    public static void ConvertAndFixPresentation(string pptSourceFileName, 
        string pptDestFileName, string pptTemplateFileName, bool appWindowVisible)
    {
        if (KillAllProcesses("POWERPNT"))
            Console.WriteLine("MS PowerPoint (POWERPNT.EXE) is still running -> process terminated.");

        MsoTriState pptAppWindowsVisible = appWindowVisible ?
            MsoTriState.msoTrue : MsoTriState.msoFalse;
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

            FixCodeBoxes(pptSource, pptDestination);

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
            if (!appWindowVisible)
                pptDestination.Close();
        }
        finally
        {
            if (!appWindowVisible)
            {
                // Quit the MS PowerPoint application
                pptApp.Quit();

                // Release any associated .NET proxies for the COM objects, which are not in use
                // Intentionally we call the garbace collector twice
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
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
                if (shape.Type == MsoShapeType.msoPlaceholder)
                    if (shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderTitle)
                        if (shape.HasTextFrame == MsoTriState.msoTrue)
                            slideTitleShape = shape;
            }
            if (slideTitleShape == null)
            {
                if (slide.Shapes.Placeholders.Count > 0)
                {
                    Shape firstShape = slide.Shapes.Placeholders[1];
                    if (firstShape?.HasTextFrame == MsoTriState.msoTrue)
                        slideTitleShape = firstShape;
                }
            }
            slideTitleShapes.Add(slideTitleShape);

            if (includeSubtitles)
            {
                // Extract also subtitles
                foreach (Shape shape in slide.Shapes.Placeholders)
                {
                    if (shape.Type == MsoShapeType.msoPlaceholder)
                        if (shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderSubtitle)
                            if (shape.HasTextFrame == MsoTriState.msoTrue)
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
		for (int i = presentation.Slides.Count; i > 0; i--)
			presentation.Slides[i].Delete();
	}

    static void CopySlidesAndSections(Presentation pptSource, Presentation pptDestination)
    {
        Console.WriteLine("Copying all slides and sections from the source presentation...");

        // Copy all slides from the source presentation
        Console.WriteLine("  Copying all slides from the source presentation...");
        pptDestination.Slides.InsertFromFile(pptSource.FullName, 0);

        // Copy all sections from the source presentation
        Console.WriteLine("  Copying all sections from the source presentation...");
        int sectionSlideIndex = 1;
        for (int sectNum = 1; sectNum <= pptSource.SectionProperties.Count; sectNum++)
        {
            string sectionName = pptSource.SectionProperties.Name(sectNum);
            sectionName = FixEnglishTitleCharacterCasing(sectionName);
            if (sectionSlideIndex <= pptDestination.Slides.Count)
                pptDestination.SectionProperties.AddBeforeSlide(sectionSlideIndex, sectionName);
            else
                pptDestination.SectionProperties.AddSection(sectNum, sectionName);
            sectionSlideIndex += pptSource.SectionProperties.SlidesCount(sectNum);
        }
    }

    static void FixCodeBoxes(Presentation pptSource, Presentation pptDestination)
    {
        Console.WriteLine("Fixing source code boxes...");

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
                    if (newShape.HasTextFrame == MsoTriState.msoTrue &&
                        newShape.TextFrame.HasText == MsoTriState.msoTrue)
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
                            MsoLanguageID.msoLanguageIDEnglishUS;
                        // newShape.TextFrame.TextRange.NoProofing = MsoTriState.msoTrue;
                    }
                }

                Console.WriteLine($"  Fixed the code box styling at slide #{slideNum}");
            }
        }
    }

    static void CopyDocumentProperties(Presentation pptSource, Presentation pptDestination)
    {
        Console.WriteLine("Copying document properties (metadata)...");

        object srcDocProperties = pptSource.BuiltInDocumentProperties;
        string title = GetObjectProperty(srcDocProperties, "Title")
            ?.ToString()?.Replace(',', ';');
        string subject = GetObjectProperty(srcDocProperties, "Subject")
            ?.ToString()?.Replace(',', ';');
        string category = GetObjectProperty(srcDocProperties, "Category")
            ?.ToString()?.Replace(',', ';');
        string keywords = GetObjectProperty(srcDocProperties, "Keywords")
            ?.ToString()?.Replace(',', ';');

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

    static void FixInvalidSlideLayouts(Presentation presentation)
    {
        Console.WriteLine("Fixing the invalid slide layouts...");

        var layoutMappings = new Dictionary<string, string> {
            { "Presentation Title Slide", "Presentation Title Slide" },
            { "Section Title Slide", "Section Title Slide" },
            { "Section Slide", "Section Title Slide" },
            { "Title Slide", "Section Title Slide" },
            { "Заглавен слайд", "Section Title Slide" },
            { "Demo Slide", "Demo Slide" },
            { "Live Exercise Slide", "Section Title Slide" },
            { "Background Slide", "Section Title Slide" },
            { "Important Concept", "Important Concept" },
            { "Important Example", "Important Example" },
            { "Table of Content", "Table of Contents" },
            { "Table of Contents", "Table of Contents" },
            { "Comparison Slide", "Comparison Slide" },
            { "Title and Content", "Title and Content" },
            { "Заглавие и съдържание", "Title and Content" },
            { "Source Code Example", "Source Code Example" },
            { "Image and Content", "Image and Content" },
            { "Questions Slide", "Questions Slide" },
            { "Blank Slide", "Blank Slide" },
            { "Block scheme", "Blank Slide" },
            { "About Slide", "About Slide" },
            { "Last", "About Slide" },
            { "Comparison Slide Dark", "Comparison Slide" },
            { "", "" },
        };
        // Add layout names like "1_Section Title Slide", "2_Demo Slide"
        foreach (string layoutName in layoutMappings.Keys.ToArray())
        {
            var mappedName = layoutMappings[layoutName];
            for (int i = 1; i < 99; i++)
                layoutMappings["" + i + "_" + layoutName] = mappedName;
        }

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
                if (customLayoutsByName.ContainsKey(newLayoutName))
                {
                    slide.CustomLayout = customLayoutsByName[newLayoutName];
                }
                else
                {
                    slide.CustomLayout = customLayoutsByName[defaultLayoutName];
                    Console.WriteLine($"  Cannot find layout [{newLayoutName}] --> using [{defaultLayoutName}] instead");
                }
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
        Console.WriteLine("Replacing incorrect hyperlinks in the slides...");

        var hyperlinksToReplace = new Dictionary<string, string> {
            { "http://softuni.bg", "https://softuni.bg" },
            { "http://softuni.foundation/", "https://softuni.foundation" },
        };

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
        Console.WriteLine("Fixing section title slides...");

        var sectionTitleSlides = presentation.Slides.Cast<Slide>()
            .Where(slide => slide.CustomLayout.Name == "Section Title Slide");

		foreach (Slide slide in sectionTitleSlides)
        {
            // Collect non-empty text shapes from the slide (starting from title and subtitle)
            List<Shape> textShapes = new List<Shape>();

            var titleShapes = FindNonEmptyTextShapesByType(slide, PpPlaceholderType.ppPlaceholderTitle);
            textShapes.AddRange(titleShapes);

			var titleCenterShapes = FindNonEmptyTextShapesByType(slide, PpPlaceholderType.ppPlaceholderCenterTitle);
			textShapes.AddRange(titleCenterShapes);

			var subtitleShapes = FindNonEmptyTextShapesByType(slide, PpPlaceholderType.ppPlaceholderSubtitle);
			textShapes.AddRange(subtitleShapes);

			var bodyShapes = FindNonEmptyTextShapesByType(slide, PpPlaceholderType.ppPlaceholderBody);
			textShapes.AddRange(bodyShapes);

			// Extract the texts from all text shapes
            List<string> slideTexts = new List<string>();
            foreach (Shape shape in textShapes)
            {
				slideTexts.Add(shape.TextFrame.TextRange.Text);
			}

			// Delete all collected text shapes --> shapes from the placeholders will appear instead
			foreach (var shape in textShapes)
            {
				shape.Delete();
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

            Console.WriteLine($" Fixed section slide #{slide.SlideNumber}: {slideTexts.FirstOrDefault()}");
        }

		List<Shape> FindNonEmptyTextShapesByType(Slide slide, PpPlaceholderType placeholderType)
		{
            List<Shape> shapes = new List<Shape>();
			foreach (Shape shape in slide.Shapes)
			{
				try
				{
					if (shape.HasTextFrame == MsoTriState.msoTrue
						&& shape.PlaceholderFormat.Type == placeholderType
						&& shape.TextFrame.TextRange.Text != "")
					{
                        shapes.Add(shape);
					}
				}
				catch (Exception)
				{
					// Silently ignore --> the shape is not a placeholder
				}
			}

            return shapes;
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
            string newTitle = FixEnglishTitleCharacterCasing(slideTitles[i]);
            if (newTitle != null && titleMappings.ContainsKey(newTitle))
                newTitle = titleMappings[newTitle];
            if (newTitle != slideTitles[i])
            {
                Console.WriteLine($"  Replaced slide #{i} title: [{slideTitles[i]}] -> [{newTitle}]");
                slideTitleShapes[i].TextFrame.TextRange.Text = newTitle;
            }
        }
    }

    static Language DetectPresentationLanguage(Presentation presentation)
    {
        Console.WriteLine("Detecting presentation language...");

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

        Language lang = (bulgarianLettersCount > englishLettersCount / 2) ?
            Language.BG : Language.EN;
        Console.WriteLine($"  Language detected: {lang}");
        return lang;
    }

    static void FixQuestionsSlide(Presentation presentation, string pptTemplateFileName,
        List<string> pptTemplateSlideTitles, Language lang)
    {
        Console.WriteLine("Fixing the [Questions] slide...");

        int questionsSlideIndexInTemplate =
            pptTemplateSlideTitles.LastIndexOf("Questions?");
        if (lang == Language.BG || questionsSlideIndexInTemplate == -1) 
        {
            int questionsSlideIndexBG = pptTemplateSlideTitles.LastIndexOf("Въпроси?");
            if (questionsSlideIndexBG != -1)
                questionsSlideIndexInTemplate = questionsSlideIndexBG;
        }

        if (questionsSlideIndexInTemplate == -1)
        {
            Console.WriteLine($"  Cannot find the [Questions] slide --> operation skipped");
            return;
        }

        for (int slideNum = 1; slideNum <= presentation.Slides.Count; slideNum++)
        {
            Slide slide = presentation.Slides[slideNum];
            if (slide.CustomLayout.Name == "Questions Slide" ||
                slide.CustomLayout.Name == "1_Questions Slide")
            {
                Console.WriteLine($"  Found the [Questions] slide #{slideNum} --> replacing it from the template");

                presentation.Slides[slideNum].Delete();
                presentation.Slides.InsertFromFile(pptTemplateFileName, slideNum - 1,
                    questionsSlideIndexInTemplate + 1, questionsSlideIndexInTemplate + 1);
            }
        }
    }

    static void FixLicenseSlide(Presentation presentation, string pptTemplateFileName,
        List<string> pptTemplateSlideTitles, Language lang)
    {
        Console.WriteLine("Fixing the [License] slide...");

        int licenseSlideIndexInTemplate = pptTemplateSlideTitles.LastIndexOf("License");
        if (lang == Language.BG || licenseSlideIndexInTemplate == -1)
        {
            int licenseSlideIndexBG = pptTemplateSlideTitles.LastIndexOf("Лиценз");
            if (licenseSlideIndexBG != -1)
                licenseSlideIndexInTemplate = licenseSlideIndexBG;
        }

        if (licenseSlideIndexInTemplate == -1)
        {
            Console.WriteLine($"  Cannot find the [License] slide --> operation skipped");
            return;
        }

        var slideTitles = ExtractSlideTitles(presentation);
        for (int slideNum = 1; slideNum <= presentation.Slides.Count; slideNum++)
        {
            if (slideTitles[slideNum - 1] == "License" ||
                slideTitles[slideNum - 1] == "Лиценз")
            {
                Console.WriteLine($"  Found the [License] slide #{slideNum} --> replacing it from the template");

                presentation.Slides[slideNum].Delete();
                presentation.Slides.InsertFromFile(pptTemplateFileName, slideNum - 1,
                    licenseSlideIndexInTemplate + 1, licenseSlideIndexInTemplate + 1);
            }
        }
    }

    static void FixAboutSoftUniSlide(Presentation presentation, string pptTemplateFileName,
        List<string> pptTemplateSlideTitles, Language lang)
    {
        Console.WriteLine("Fixing the [About] slide...");

        var aboutSlidePossibleTitles = new List<(string, Language)>()
        {
            ("Trainings @ Software University (SoftUni)", Language.EN),
            ("About SoftUni Digital", Language.EN),
            ("About SoftUni Creative", Language.EN),
            ("About SoftUni Kids", Language.EN),
            ("Обучения в СофтУни", Language.BG),
            ("Обучения в Софтуерен университет (СофтУни)", Language.BG),
        };

        // Find the replacement slide index from the template
        // for the [About] slide in the current language 
        int aboutSlideReplacementIndexInTemplate = -1;
        foreach (var title in aboutSlidePossibleTitles.Where(t => t.Item2 == lang))
            aboutSlideReplacementIndexInTemplate = Math.Max(
                aboutSlideReplacementIndexInTemplate,
                pptTemplateSlideTitles.LastIndexOf(title.Item1));
        
        // Remove the language filter if the [About] slide is not found
        if (aboutSlideReplacementIndexInTemplate == -1)
            foreach (var title in aboutSlidePossibleTitles)
                aboutSlideReplacementIndexInTemplate = Math.Max(
                    aboutSlideReplacementIndexInTemplate,
                    pptTemplateSlideTitles.LastIndexOf(title.Item1));

        if (aboutSlideReplacementIndexInTemplate == -1)
        {
            Console.WriteLine($"  Cannot find the [About] slide --> operation skipped");
            return;
        }

        // Replace the [About] slides from the presentation
        // with the [About] slide from the template
        HashSet<string> slideTitlesToReplace = new HashSet<string>(
            aboutSlidePossibleTitles.Select(t => t.Item1));
        var slideTitles = ExtractSlideTitles(presentation);
        for (int slideNum = 1; slideNum <= presentation.Slides.Count; slideNum++)
        {
            if (slideTitlesToReplace.Contains(slideTitles[slideNum - 1]))
            {
                Console.WriteLine($"  Found the [About] slide #{slideNum} --> replacing it from the template");
                presentation.Slides[slideNum].Delete();
                presentation.Slides.InsertFromFile(pptTemplateFileName, slideNum - 1,
                    aboutSlideReplacementIndexInTemplate + 1, aboutSlideReplacementIndexInTemplate + 1);
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
                if (shape.Type == MsoShapeType.msoPlaceholder)
                    if (shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderSlideNumber)
                        return shape;
            return null;
        }

        Console.Write("Fixing slide numbering...");

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
                bool isSlideNumberTextBox = shape.Type == MsoShapeType.msoTextBox
                    && shape.Name.Contains("Slide Number");
                bool isSlideNumberPlaceholder = shape.Type == MsoShapeType.msoPlaceholder
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

            Console.Write("."); // Display progress of the current operation
        }
        Console.WriteLine();
    }

    static void FixSlideNotesPages(Presentation presentation)
    {
        Shape FindNotesFooter()
        {
            var footerShape = presentation.NotesMaster.Shapes.OfType<Shape>().Where(
                shape => shape.Type == MsoShapeType.msoPlaceholder
                && shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderFooter)
                .FirstOrDefault();
            return footerShape;
        }

        Console.WriteLine("Fixing slide notes pages...");

        Shape footerFromNotesMaster = FindNotesFooter();
        if (footerFromNotesMaster != null)
        {
			footerFromNotesMaster.Copy();

			for (int slideNum = 1; slideNum <= presentation.Slides.Count; slideNum++)
			{
				Slide slide = presentation.Slides[slideNum];
				if (slide.HasNotesPage == MsoTriState.msoTrue)
				{
					var slideNotesFooter = slide.NotesPage.Shapes.OfType<Shape>().Where(
						shape => shape.Type == MsoShapeType.msoPlaceholder
						&& shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderFooter)
						.FirstOrDefault();
					if (slideNotesFooter != null)
						slideNotesFooter.Delete();
					slide.NotesPage.Shapes.Paste();
				}
			}
		}
    }
}
