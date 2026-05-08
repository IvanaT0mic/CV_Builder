using CVBuilder.Dtos;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CVBuilder.Services
{
    public interface ICvBuilderService
    {
        Task<byte[]> GenerateCVAsync(CVFieldValuesDto data);
    }

    public class CvBuilderService : ICvBuilderService
    {
        public async Task<byte[]> GenerateCVAsync(CVFieldValuesDto data)
        {
            var templatePath = Path.Combine(
                AppContext.BaseDirectory,
                "Templates",
                "CV_Template.docx");

            if (!File.Exists(templatePath))
                throw new FileNotFoundException("CV template not found.");

            using var ms = new MemoryStream();

            await using (var fs = File.OpenRead(templatePath))
            {
                await fs.CopyToAsync(ms);
            }

            ms.Position = 0;

            var replacements = new Dictionary<string, string>
            {
                ["I_FIRST_NAME"] = data.FirstName ?? "",
                ["I_LAST_NAME"] = data.LastName ?? "",
                ["I_TITLE"] = data.Title ?? "",
                ["I_ABOUT_ME_DESCRIPTION"] = data.AboutMe ?? "",
            };

            using (var wordDoc = WordprocessingDocument.Open(ms, true))
            {
                var body = wordDoc.MainDocumentPart!.Document.Body!;

                GenerateExperiences(body, data.Experiences);

                foreach (var text in body.Descendants<Text>().ToList())
                {
                    if (text.Text.Contains("I_SOCIAL_NAME"))
                    {
                        var parentParagraph = text.Ancestors<Paragraph>().FirstOrDefault();
                        if (parentParagraph != null && data.Socials?.Any() == true)
                        {
                            var dataList = data.Socials.Select(x => (x.Platform, x.Url)).ToList();
                            // Platforma bold crna, Link normal zelen ("3CD52E")
                            var newRows = CreateBulletItems(dataList, parentParagraph, "000000", true, "3CD52E", false);
                            InsertAndRemove(body, parentParagraph, newRows);
                            continue;
                        }
                    }

                    if (text.Text.Contains("I_TRAINING_NAME"))
                    {
                        var parentParagraph = text.Ancestors<Paragraph>().FirstOrDefault();
                        if (parentParagraph != null && data.Trainings?.Any() == true)
                        {
                            var dataList = data.Trainings.Select(x => (x.TrainingName, x.TrainingDate)).ToList();
                            // Naziv treninga bold crn, Datum normal siv ("808080")
                            var newRows = CreateBulletItems(dataList, parentParagraph, "000000", true, "808080", false);
                            InsertAndRemove(body, parentParagraph, newRows);
                            continue;
                        }
                    }

                    if (text.Text.Contains("I_DEGREE_INFO"))
                    {
                        var parentParagraph = text.Ancestors<Paragraph>().FirstOrDefault();
                        if (parentParagraph != null && data.Educations?.Any() == true)
                        {
                            var dataList = data.Educations.Select(x => (x.EducationName, x.EduEndDate ?? "")).ToList();
                            // Edukacija bold crna, Datum normal siv ("808080")
                            var newRows = CreateBulletItems(dataList, parentParagraph, "000000", true, "808080", false);
                            InsertAndRemove(body, parentParagraph, newRows);
                            continue;
                        }
                    }

                    if (text.Text.Contains("I_PROFILE_GRADES"))
                    {
                        var parentRun = text.Parent as Run;
                        var parentParagraph = parentRun?.Parent;

                        if (parentParagraph != null && data.ProfileGrades != null)
                        {
                            var originalStyle = parentRun.RunProperties;
                            var formattedElements = CreateProfileWithGrades(data.ProfileGrades, "3CD52E", originalStyle);

                            foreach (var element in formattedElements)
                            {
                                parentParagraph.InsertBefore(element, parentRun);
                            }

                            parentRun.Remove();
                            continue;
                        }
                    }

                    if (text.Text.Contains("I_TOOL_STACK"))
                    {
                        var parentRun = text.Parent as Run;
                        var parentParagraph = parentRun?.Parent;

                        if (parentParagraph != null)
                        {
                            var originalStyle = parentRun.RunProperties;

                            var formattedElements = CreateFormattedList(data.ToolStack ?? "", "3CD52E", originalStyle);

                            foreach (var element in formattedElements)
                            {
                                parentParagraph.InsertBefore(element, parentRun);
                            }

                            parentRun.Remove();
                            continue;
                        }
                    }

                    if (text.Text.Contains("I_SKILLS"))
                    {
                        var parentRun = text.Parent as Run;
                        var parentParagraph = parentRun?.Parent;

                        if (parentParagraph != null)
                        {
                            var originalStyle = parentRun.RunProperties;

                            var formattedElements = CreateFormattedList(data.Skills ?? "", "3CD52E", originalStyle);

                            foreach (var element in formattedElements)
                            {
                                parentParagraph.InsertBefore(element, parentRun);
                            }

                            parentRun.Remove();
                            continue;
                        }
                    }

                    if (text.Text.Contains("I_LANGUAGE"))
                    {
                        var parentRun = text.Parent as Run;
                        var parentParagraph = parentRun?.Parent;

                        if (parentParagraph != null)
                        {
                            var originalStyle = parentRun.RunProperties;

                            var formattedElements = CreateFormattedList(data.Languages ?? "", "3CD52E", originalStyle);

                            foreach (var element in formattedElements)
                            {
                                parentParagraph.InsertBefore(element, parentRun);
                            }

                            parentRun.Remove();
                            continue;
                        }
                    }

                    foreach (var (placeholder, value) in replacements)
                    {
                        if (text.Text.Contains(placeholder))
                        {
                            text.Text = text.Text.Replace(placeholder, value);
                        }
                    }
                }

                foreach (var headerPart in wordDoc.MainDocumentPart.HeaderParts)
                {
                    ReplaceInPart(headerPart.Header.Descendants<Text>(), replacements);
                }

                foreach (var footerPart in wordDoc.MainDocumentPart.FooterParts)
                {
                    ReplaceInPart(footerPart.Footer.Descendants<Text>(), replacements);
                }

                wordDoc.MainDocumentPart.Document.Save();
            }

            return ms.ToArray();
        }

        private IEnumerable<Paragraph> CreateBulletItems(
            List<(string Part1, string Part2)> items,
            Paragraph templateParagraph,
            string color1, bool bold1,
            string color2, bool bold2)
        {
            var newParagraphs = new List<Paragraph>();

            // Pokušavamo da izvučemo originalni stil iz prvog dostupnog Run-a u templejtu
            var originalRunProps = templateParagraph.Descendants<RunProperties>().FirstOrDefault();

            foreach (var item in items)
            {
                var newPara = (Paragraph)templateParagraph.CloneNode(true);
                newPara.RemoveAllChildren<Run>();

                // PRVI DEO
                var run1 = new Run();
                var rPr1 = originalRunProps != null ? (RunProperties)originalRunProps.CloneNode(true) : new RunProperties();
                rPr1.Bold = bold1 ? new Bold() : new Bold { Val = false };
                rPr1.Color = new Color { Val = color1 };
                run1.AppendChild(rPr1);
                run1.AppendChild(new Text(item.Part1) { Space = SpaceProcessingModeValues.Preserve });

                // TAB
                newPara.AppendChild(run1);
                newPara.AppendChild(new Run(new TabChar()));

                // DRUGI DEO
                var run2 = new Run();
                var rPr2 = originalRunProps != null ? (RunProperties)originalRunProps.CloneNode(true) : new RunProperties();
                rPr2.Bold = bold2 ? new Bold() : new Bold { Val = false };
                rPr2.Color = new Color { Val = color2 };
                run2.AppendChild(rPr2);
                run2.AppendChild(new Text(item.Part2) { Space = SpaceProcessingModeValues.Preserve });

                newPara.AppendChild(run2);

                newParagraphs.Add(newPara);
            }
            return newParagraphs;
        }

        private IEnumerable<OpenXmlElement> CreateProfileWithGrades(List<ProfileGradesDto> grades, string colorHex, RunProperties? originalProperties)
        {
            var elements = new List<OpenXmlElement>();

            for (int i = 0; i < grades.Count; i++)
            {
                var item = grades[i];

                var textRun = new Run();
                if (originalProperties != null) textRun.AppendChild(originalProperties.CloneNode(true));
                textRun.AppendChild(new Text(item.Skill + " ") { Space = SpaceProcessingModeValues.Preserve });
                elements.Add(textRun);

                item.Grade = item.Grade ?? 0;
                for (int g = 0; g < item.Grade; g++)
                {
                    var symbolRun = new Run();
                    if (originalProperties != null) symbolRun.AppendChild(originalProperties.CloneNode(true));

                    var srp = symbolRun.GetFirstChild<RunProperties>() ?? symbolRun.PrependChild(new RunProperties());
                    srp.RunFonts = new RunFonts { Ascii = "Wingdings", HighAnsi = "Wingdings" };
                    srp.Color = new Color { Val = colorHex };

                    symbolRun.AppendChild(new Text("\uF0A7") { Space = SpaceProcessingModeValues.Preserve });
                    elements.Add(symbolRun);
                }

                if (i < grades.Count - 1)
                {
                    var breakRun = new Run();
                    if (originalProperties != null) breakRun.AppendChild(originalProperties.CloneNode(true));
                    breakRun.AppendChild(new Break());
                    elements.Add(breakRun);
                }
            }

            return elements;
        }

        private IEnumerable<OpenXmlElement> CreateFormattedList(string input, string colorHex, RunProperties? originalProperties)
        {
            var elements = new List<OpenXmlElement>();
            var items = input.Split(',', StringSplitOptions.RemoveEmptyEntries);

            foreach (var item in items)
            {
                var symbolRun = new Run();

                if (originalProperties != null)
                    symbolRun.AppendChild(originalProperties.CloneNode(true));

                var srp = symbolRun.GetFirstChild<RunProperties>() ?? symbolRun.PrependChild(new RunProperties());
                srp.RunFonts = new RunFonts { Ascii = "Wingdings", HighAnsi = "Wingdings" };
                srp.Color = new Color { Val = colorHex };

                symbolRun.AppendChild(new Text("\uF0A7") { Space = SpaceProcessingModeValues.Preserve });

                var textRun = new Run();

                if (originalProperties != null)
                    textRun.AppendChild(originalProperties.CloneNode(true));

                textRun.AppendChild(new Text(" " + item.Trim() + "  ") { Space = SpaceProcessingModeValues.Preserve });

                elements.Add(symbolRun);
                elements.Add(textRun);
            }

            return elements;
        }

        private void GenerateExperiences(Body body, List<ExperienceDto> experiences)
        {
            // 1. Pronađi markere u dokumentu
            var startPara = body.Descendants<Paragraph>().FirstOrDefault(p => p.InnerText.Contains("I_EXP_START"));
            var endPara = body.Descendants<Paragraph>().FirstOrDefault(p => p.InnerText.Contains("I_EXP_END"));

            if (startPara == null || endPara == null) return;

            // 2. "Usisaj" sve paragrafe između markera u listu (to je naš šablon)
            var templateParagraphs = new List<Paragraph>();
            var current = startPara.NextSibling();
            while (current != null && current != endPara)
            {
                if (current is Paragraph p)
                {
                    templateParagraphs.Add((Paragraph)p.CloneNode(true));
                }
                current = current.NextSibling();
            }

            // 3. Generiši nove blokove za svako iskustvo iz DTO-a
            var lastElement = endPara; // Ubacivaćemo ISPOD kraja markera

            foreach (var exp in experiences)
            {
                foreach (var tempPara in templateParagraphs)
                {
                    var newPara = (Paragraph)tempPara.CloneNode(true);

                    // -- A: Rukovanje običnim poljima --
                    var texts = newPara.Descendants<Text>().ToList();

                    // Provera za Responsibilities bullet (specijalni slučaj unutar bloka)
                    if (newPara.InnerText.Contains("I_EXP_RESPONSIBILITIES"))
                    {
                        var respParent = lastElement.Parent;
                        foreach (var resp in exp.ExpResponsibilities ?? new List<string>())
                        {
                            var bulletPara = (Paragraph)tempPara.CloneNode(true);
                            foreach (var t in bulletPara.Descendants<Text>())
                                t.Text = t.Text.Replace("I_EXP_RESPONSIBILITIES", resp);

                            respParent.InsertAfter(bulletPara, lastElement);
                            lastElement = bulletPara;
                        }
                        continue; // Preskačemo dodavanje originalnog template reda za responsibilities
                    }

                    // Regularna zamena teksta u ostalim redovima paragrafa
                    foreach (var t in texts)
                    {
                        t.Text = t.Text.Replace("I_EXP_POSITION", exp.ExpPosition ?? "")
                                       .Replace("I_START_DATE", exp.ExpStartDate ?? "")
                                       .Replace("I_END_DATE", exp.ExpEndDate ?? "Present")
                                       .Replace("I_COMPANY_NAME", exp.ExpCompanyName ?? "")
                                       .Replace("I_PROJECT_NAME", exp.ExpProjectName ?? "")
                                       .Replace("I_PROJECT_DESCRIPTION", exp.ExpProjectDescription ?? "")
                                       .Replace("I_EXP_TECH_STACK", exp.ExpTechStack ?? "");
                    }

                    lastElement.Parent.InsertAfter(newPara, lastElement);
                    lastElement = newPara;
                }

                // -- B: Ubacivanje separatora (onaj red sa linijom/borderom) --
                // Kloniramo endPara jer on nosi vizuelni stil kraja bloka (tvoju misterioznu crtu)
                var separator = (Paragraph)endPara.CloneNode(true);
                foreach (var t in separator.Descendants<Text>()) t.Text = ""; // Brišemo tekst "I_EXP_END"

                lastElement.Parent.InsertAfter(separator, lastElement);
                lastElement = separator;
            }

            // 4. Čišćenje: Obriši originalne template paragrafe i markere
            var toDelete = new List<OpenXmlElement>();
            var pointer = startPara.NextSibling();
            while (pointer != null && pointer != endPara)
            {
                toDelete.Add(pointer);
                pointer = pointer.NextSibling();
            }

            foreach (var el in toDelete) el.Remove();
            startPara.Remove();
            endPara.Remove();
        }

        private void InsertAndRemove(Body body, Paragraph templateParagraph, IEnumerable<Paragraph> newItems)
        {
            var lastInserted = templateParagraph;
            foreach (var item in newItems)
            {
                lastInserted = body.InsertAfter(item, lastInserted);
            }
            templateParagraph.Remove();
        }

        private static void ReplaceInPart(
            IEnumerable<Text> texts,
            Dictionary<string, string> replacements)
        {
            foreach (var text in texts)
            {
                foreach (var (k, v) in replacements)
                {
                    if (text.Text.Contains(k))
                    {
                        text.Text = text.Text.Replace(k, v);
                    }
                }
            }
        }
    }
}