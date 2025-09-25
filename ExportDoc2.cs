using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Sitecore;
using Sitecore.Data.Items;
using Sitecore.Data.Fields;
using Sitecore.Diagnostics;
using Sitecore.Text;
using Sitecore.Web.UI.Sheer;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using HtmlAgilityPack;
using NotesFor.HtmlToOpenXml;
using ViridianSpark.CourseCatalog;
using Sitecore.Data;
using SmartCatalog.SCMiscellaneous;

namespace SmartCatalog.ExportDocx2
{
    public class paragraphPart
    {
        public string name { get; set; }
        public string text { get; set; }
        public string style { get; set; }
        public string requires { get; set; }
    }

    public class courseParagraph
    {
        public string style { get; set; }
        public string requires { get; set; }
        public List<paragraphPart> parts = new List<paragraphPart>();
        public bool isHTML { get; set; }
        public bool isLink { get; set; }
        public string name { get; set; }
        public bool reqsMet { get; set; }
    }

    public class Job
    {
        public string wordTemplatePath { get; set; }
        public string wordOutputPath { get; set; }
        public string configFilePath { get; set; }
        public string configFileField { get; set; }

        public string calcSubtotals { get; set; }
        public string calcTotals { get; set; }
        public string includeCredits { get; set; }
        public string calcDegReq { get; set; }

        public string col1Width { get; set; }
        public string col2Width { get; set; }
        public string col3Width { get; set; }
        public string col4Width { get; set; }
        public string col5Width { get; set; }

        public int topMargin { get; set; }
        public int bottomMargin { get; set; }
        public int defaultNumberOfColumns { get; set; }
        public uint pageHeight { get; set; }
        public uint pageWidth { get; set; }
        public uint rightMargin { get; set; }
        public uint leftMargin { get; set; }
        public uint footerDistance { get; set; }
        public uint headerDistance { get; set; }
        public uint spaceBetweenColumns { get; set; }

        public int bmIdCounter { get; set; }
        public int level { get; set; }

        public bool hasConfigFile { get; set; }
        public XDocument configTree { get; set; }

        public Item startItem { get; set; }
        public Item institutionItem { get; set; }
        public Item settingsItem { get; set; }
        public Item yearItem { get; set; }
        public Item catalogItem { get; set; }

        public Item programItem { get; set; }

        public string yearName { get; set; }
        public string catalogTitle { get; set; }
        public string instTitle { get; set; }
        public string headerText { get; set; }

        public List<string> bookmarkables { get; set; }

        public List<string> errorNames { get; set; }
        public string errorsMessage { get; set; }

        public MemoryStream mem { get; set; }
        public WordprocessingDocument doc { get; set; }
        public Body body { get; set; }
        public Object[] args { get; set; }

        public Database database { get; set; }

        public dynamic programInserter { get; set; }

        private string prependBaseWebsitePath(string relativePath)
        {
            if (string.IsNullOrWhiteSpace(relativePath)) return relativePath;
            string baseWebsitePath = Sitecore.Configuration.Settings.GetSetting("WebsiteLocation").TrimEnd('\\');
            string trimmed = relativePath.TrimStart('\\');
            return baseWebsitePath + "\\" + trimmed;
        }

        private static string GetString(Item item, string fieldName)
        {
            return item?.Fields?[fieldName]?.Value ?? string.Empty;
        }

        private static bool IsChecked(Item item, string fieldName)
        {
            return string.Equals(GetString(item, fieldName), "1", StringComparison.Ordinal);
        }

        private static string NormalizeCredits(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;
            string cleaned = text.Replace(" credit hours", "")
                                 .Replace(" credit hour", "")
                                 .Replace(" credits", "")
                                 .Replace(" credit", "")
                                 .Replace(" units", "")
                                 .Replace(" unit", "");
            return cleaned == "0" ? string.Empty : cleaned;
        }

        private static string NormalizeOfferedString(string offered)
        {
            if (string.IsNullOrWhiteSpace(offered)) return string.Empty;
            string s = offered;
            s = s.Replace("Fall", "F").Replace("Summer", "Su").Replace("Spring", "Sp");
            s = s.Replace("fall", "F").Replace("summer", "Su").Replace("spring", "Sp");
            s = s.Replace("Offered ", "");
            s = s.Replace(".", "");
            s = s.Replace(" and ", ", ");
            if (s.Length > 1)
            {
                s = char.ToUpper(s[0]) + s.Substring(1);
            }
            return s;
        }

        public Job(Item _startItem, string databaseName)
        {
            database = string.IsNullOrEmpty(databaseName)
                ? Sitecore.Configuration.Factory.GetDatabase("master")
                : Sitecore.Configuration.Factory.GetDatabase(databaseName);

            bookmarkables = new List<string>();
            bmIdCounter = 0;
            errorNames = new List<string>();

            startItem = _startItem;
            institutionItem = startItem?.Axes?.SelectSingleItem("ancestor-or-self::*[@@templatename='Institution']");

            hasConfigFile = false;
            configTree = new XDocument();

            settingsItem = institutionItem?.Children?.FirstOrDefault(i => i.TemplateName == "Settings Folder");
            if (settingsItem == null)
            {
                Log.Error("ExportDocx2: User " + Context.User.Name + " couldn't access settings folder, aborting docx generation.", this);
                Sitecore.Context.ClientPage.ClientResponse.Alert("Could not access a required settings item due to user permissions.\nPlease contact your institution's Curriculum Strategy administrtor to resolve this issue.");
                throw new InvalidOperationException("Settings Folder not accessible.");
            }
            settingsItem = settingsItem.Children.FirstOrDefault(i => i.TemplateName == "Print Settings");
            if (settingsItem == null)
            {
                Log.Error("ExportDocx2: User " + Context.User.Name + " couldn't access print settings, aborting docx generation.", this);
                Sitecore.Context.ClientPage.ClientResponse.Alert("Could not access a required settings item due to user permissions.\nPlease contact your institution's Curriculum Strategy administrtor to resolve this issue.");
                throw new InvalidOperationException("Print Settings not accessible.");
            }

            string typeString = "DefaultProgramInserter";
            if (!string.IsNullOrWhiteSpace(settingsItem["Custom Program Handler"]))
            {
                typeString = settingsItem["Custom Program Handler"];
            }
            string typeName = "SmartCatalog.ExportDocx2.ProgramInserters." + typeString;
            Type type = Type.GetType(typeName);
            Assert.IsNotNull(type, "Program inserter type was null. typeString = " + typeName);
            object[] programInserterArgs = new object[] { this };
            programInserter = Activator.CreateInstance(type, programInserterArgs) as ProgramInserters.ProgramInserter;

            wordTemplatePath = settingsItem["Word Template Path"];
            if (!string.IsNullOrWhiteSpace(wordTemplatePath) && wordTemplatePath.StartsWith("\\"))
            {
                wordTemplatePath = prependBaseWebsitePath(wordTemplatePath);
            }

            calcSubtotals = settingsItem["Calculate Subtotals"];
            calcTotals = settingsItem["Calculate Totals"];
            includeCredits = settingsItem["Include Credits"];
            // Fix: calcDegReq should not mirror Include Credits
            calcDegReq = settingsItem["Calculate Degree Requirements"] ?? settingsItem["Include Credits"]; // fallback if field missing

            col1Width = settingsItem["Column 1 Width"]; col2Width = settingsItem["Column 2 Width"]; col3Width = settingsItem["Column 3 Width"]; 
            col4Width = settingsItem["Column 4 Width"] ?? "0"; col5Width = settingsItem["Column 5 Width"] ?? "0";

            configFilePath = settingsItem["Config File Path"];
            configFileField = settingsItem["Custom Config"];
            if (!string.IsNullOrWhiteSpace(configFileField))
            {
                try
                {
                    configTree = XDocument.Parse(configFileField);
                    hasConfigFile = true;
                }
                catch (System.Xml.XmlException e)
                {
                    errorNames.Add("Custom Config field could not be loaded because it contains invalid XML.");
                    Log.Warn("Print: custom config XML invalid; proceeding without config.", e, this);
                    hasConfigFile = false;
                }
            }
            else
            {
                if (!string.IsNullOrWhiteSpace(configFilePath) && configFilePath.StartsWith("\\"))
                {
                    configFilePath = prependBaseWebsitePath(configFilePath);
                }
                if (!string.IsNullOrWhiteSpace(configFilePath) && File.Exists(configFilePath))
                {
                    hasConfigFile = true;
                    configTree = XDocument.Load(configFilePath, LoadOptions.PreserveWhitespace);
                }
                else
                {
                    hasConfigFile = false;
                }
            }

            try
            {
                topMargin = System.Convert.ToInt32(settingsItem["Top Margin"]);
                bottomMargin = System.Convert.ToInt32(settingsItem["Bottom Margin"]);
                pageHeight = System.Convert.ToUInt32(settingsItem["Page Height"]);
                pageWidth = System.Convert.ToUInt32(settingsItem["Page Width"]);
                rightMargin = System.Convert.ToUInt32(settingsItem["Right Margin"]);
                leftMargin = System.Convert.ToUInt32(settingsItem["Left Margin"]);
                footerDistance = System.Convert.ToUInt32(settingsItem["Footer Distance"]);
                headerDistance = System.Convert.ToUInt32(settingsItem["Header Distance"]);
                spaceBetweenColumns = System.Convert.ToUInt32(settingsItem["Space Between Columns"]);
            }
            catch
            {
                topMargin = System.Convert.ToInt32("1420");
                bottomMargin = System.Convert.ToInt32("1650");
                pageHeight = System.Convert.ToUInt32("15840");
                pageWidth = System.Convert.ToUInt32("12240");
                rightMargin = System.Convert.ToUInt32("910");
                leftMargin = System.Convert.ToUInt32("1080");
                footerDistance = System.Convert.ToUInt32("940");
                headerDistance = System.Convert.ToUInt32("720");
                spaceBetweenColumns = System.Convert.ToUInt32("720");
            }

            try
            {
                defaultNumberOfColumns = System.Convert.ToInt32(settingsItem["Default Number of Columns"]);
            }
            catch
            {
                defaultNumberOfColumns = System.Convert.ToInt32("2");
            }

            wordOutputPath = Sitecore.IO.FileUtil.GetWorkFilename(Sitecore.Configuration.Settings.TempFolderPath, startItem.Name, ".docx");

            yearItem = startItem.Axes.SelectSingleItem("ancestor-or-self::*[@@templatename='Year Folder']");
            yearName = yearItem != null ? yearItem.Name : string.Empty;

            catalogItem = startItem.Axes.SelectSingleItem("ancestor-or-self::*[@@templatename='Catalog']");
            catalogTitle = catalogItem != null ? (GetString(catalogItem, "Title").Length > 0 ? GetString(catalogItem, "Title") : catalogItem.Name) : string.Empty;

            instTitle = institutionItem != null ? (GetString(institutionItem, "Title").Length > 0 ? GetString(institutionItem, "Title") : institutionItem.Name) : string.Empty;

            headerText = " | " + instTitle + " " + catalogTitle;

            args = new Object[2];
            mem = new MemoryStream();
            errorsMessage = string.Empty;

            // Defer opening the Word document until generateDocx(), so we can ensure proper disposal
        }

        public Object[] generateDocx()
        {
            // Load template into memory stream
            byte[] byteArray = File.ReadAllBytes(wordTemplatePath);
            mem.SetLength(0);
            mem.Write(byteArray, 0, byteArray.Length);
            mem.Position = 0;

            // Collect bookmarkable IDs (consider optimizing if needed)
            Item[] descendantItems = startItem.Axes.GetDescendants();
            foreach (Item descendant in descendantItems)
            {
                bookmarkables.Add(descendant.ID.ToString().Replace("-", "").Replace("{", "").Replace("}", ""));
            }

            var errorsBuilder = new StringBuilder();

            using (doc = WordprocessingDocument.Open(mem, true))
            {
                body = doc.MainDocumentPart.Document.Body;

                // Replace header placeholder across all headers to avoid hard-coded rIds
                const string placeholder1 = " | EvenPageHEader"; // original casing in template
                const string placeholder2 = " | EvenPageHeader"; // corrected casing variant
                foreach (var headerPart in doc.MainDocumentPart.HeaderParts)
                {
                    var header = headerPart.Header;
                    if (header == null) continue;
                    foreach (Text t in header.Descendants<Text>())
                    {
                        if (t.Text != null && (t.Text.Contains(placeholder1) || t.Text.Contains(placeholder2)))
                        {
                            t.Text = t.Text.Replace(placeholder1, headerText).Replace(placeholder2, headerText);
                        }
                    }
                }

                int level = 0;
                Item currentItem = startItem;
                while (currentItem != null && !(currentItem.TemplateName == "Catalog"))
                {
                    currentItem = currentItem.Parent;
                    level++;
                }

                body.RemoveAllChildren();
                addToc();
                addToBody(startItem);
                addIndex();
            }

            // Build error HTML once
            errorsBuilder.Append("<ul>");
            foreach (string msg in errorNames)
            {
                errorsBuilder.Append("<li>").Append(msg).Append("</li>\n");
            }
            errorsBuilder.Append("</ul>");
            errorsMessage = errorsBuilder.ToString();
            Log.Warn("Print: HTML processing errors: " + errorsMessage, this);
            errorNames.Clear();

            mem.Position = 0;
            args[0] = mem;
            args[1] = errorsMessage;
            return args;
        }

        public void addToBody(Sitecore.Data.Items.Item item)
        {
            try
            {
                switch (item.TemplateName)
                {
                    case "Content Section":
                    case "Handbook Content Section":
                        addContent(item);
                        break;
                    case "Course":
                        addCourse(item);
                        break;
                    case "Catalog":
                        addCatalog(item);
                        break;
                    case "Certificate":
                    case "Degree":
                    case "Minor":
                    case "Narrative with Course Table":
                        addAward(item);
                        break;
                    case "Division":
                        addDivision(item);
                        break;
                    case "Courses Folder":
                        addCoursesFolder(item);
                        break;
                    case "Degree Requirements":
                    case "Stevenson Minor":
                    case "Degree Requirements CM":
                    case "Degrees and Certificates Folder":
                    case "Degree-Requirements":
                    case "Requirements List":
                    case "Requirements List with Courses":
                    case "Search Results":
                    case "Program Goals":
                    case "Program-Goals":
                    case "Program Outcomes":
                    case "Program-Outcomes":
                    case "Outcomes Folder":
                    case "Course Outcomes":
                    case "Outcome":
                        break;
                    default:
                        addOther(item);
                        break;
                }
            }
            catch (Exception e)
            {
                addPara("red", "*WARNING*: Item " + item.Name + " could not be inserted", false);
                errorNames.Add(item.Paths.FullPath + "(could not be inserted)");
                Log.Warn("Print: " + item.Name + " " + item.ID.ToString() + " aborted due to error: ", e, this);
                return;
            }

            if (item.HasChildren)
            {
                level++;
                foreach (Sitecore.Data.Items.Item child in item.GetChildren())
                {
                    if (child.Fields["Include in Print"] != null)
                    {
                        if (child["Include in Print"] == "1")
                        {
                            addToBody(child);
                        }
                    }
                    else
                    {
                        addToBody(child);
                    }
                }
                level--;
            }
            if (level == 1)
            {
                insertSection(defaultNumberOfColumns, SectionMarkValues.NextPage);
            }
            return;
        }

        public void breakBefore(Item item)
        {
            if (item["Wide Section"] == "1")
            {
                insertSection(defaultNumberOfColumns, SectionMarkValues.Continuous);
            }
        }

        public void breakAfter(Item item)
        {
            if (item.Fields["Wide Section"] != null)
            {
                if (item.Fields["Wide Section"].Value == "1")
                {
                    insertSection(1, SectionMarkValues.Continuous);
                }
            }
        }

        public void insertSection(int numberOfColumns, SectionMarkValues breakType)
        {
            Paragraph paragraph1 = new Paragraph();
            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            SectionProperties sectionProperties1 = new SectionProperties();

            // Avoid hard-coded Header/Footer rIds; rely on template's existing linkage
            SectionType sectionType1 = new SectionType() { Val = breakType };
            PageSize pageSize1 = new PageSize()
            {
                Width = pageWidth,
                Height = pageHeight
            };

            PageMargin pageMargin1 = new PageMargin()
            {
                Top = topMargin,
                Right = rightMargin,
                Bottom = bottomMargin,
                Left = leftMargin,
                Header = headerDistance,
                Footer = footerDistance,
                Gutter = (UInt32Value)0U
            };

            Columns columns1 = new Columns()
            {
                Space = spaceBetweenColumns.ToString(),
                ColumnCount = System.Convert.ToInt16(numberOfColumns)
            };

            DocGrid docGrid1 = new DocGrid()
            {
                LinePitch = 360
            };

            sectionProperties1.Append(sectionType1);
            sectionProperties1.Append(pageSize1);
            sectionProperties1.Append(pageMargin1);
            sectionProperties1.Append(columns1);
            sectionProperties1.Append(docGrid1);

            paragraphProperties1.Append(sectionProperties1);
            paragraph1.Append(paragraphProperties1);
            body.Append(paragraph1);
        }

        public void addHTML(string source)
        {
            if (string.IsNullOrWhiteSpace(source))
            {
                return;
            }

            string html = source;

            // Parse once and sanitize attributes via DOM operations
            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(html);

            // Remove width/height attributes and inline styles for width/height
            var nodesWithDims = htmlDoc.DocumentNode.SelectNodes("//*[@width or @height or @style]");
            if (nodesWithDims != null)
            {
                foreach (var n in nodesWithDims)
                {
                    n.Attributes.Remove("width");
                    n.Attributes.Remove("height");
                    var style = n.GetAttributeValue("style", "");
                    if (!string.IsNullOrEmpty(style))
                    {
                        var cleaned = Regex.Replace(style, @"\b(width|height)\s*:\s*[^;]+;?", string.Empty, RegexOptions.IgnoreCase);
                        n.SetAttributeValue("style", cleaned);
                    }
                }
            }

            // Convert <li><div>.. to <li>.. and <li><p>.. to <li>..
            foreach (var li in htmlDoc.DocumentNode.SelectNodes("//li") ?? Enumerable.Empty<HtmlNode>())
            {
                if (li.FirstChild != null && (li.FirstChild.Name == "div" || li.FirstChild.Name == "p"))
                {
                    li.InnerHtml = li.FirstChild.InnerHtml;
                }
            }

            // Remove blank rows or normalize empty rows
            foreach (var tr in htmlDoc.DocumentNode.SelectNodes("//tr") ?? Enumerable.Empty<HtmlNode>())
            {
                if (string.IsNullOrWhiteSpace(tr.InnerText))
                {
                    tr.InnerHtml = "<td></td>";
                }
            }

            // Remove paragraphs inside TD by flattening
            foreach (var td in htmlDoc.DocumentNode.SelectNodes("//td") ?? Enumerable.Empty<HtmlNode>())
            {
                var ps = td.SelectNodes("p");
                if (ps == null) continue;
                string pstyle = string.Empty;
                var combined = new StringBuilder();
                int i = 0;
                foreach (var p in ps)
                {
                    if (i++ > 0) combined.Append("<br />");
                    combined.Append(p.InnerHtml);
                    pstyle = p.GetAttributeValue("style", pstyle);
                }
                td.InnerHtml = combined.ToString();
                if (!string.IsNullOrEmpty(pstyle)) td.SetAttributeValue("style", pstyle);
            }

            // Remove class from paragraphs under ordered lists
            foreach (var p in htmlDoc.DocumentNode.SelectNodes("//ol//p") ?? Enumerable.Empty<HtmlNode>())
            {
                p.SetAttributeValue("class", "");
            }

            html = htmlDoc.DocumentNode.WriteTo();

            MainDocumentPart mainPart = doc.MainDocumentPart;
            HtmlConverter converter = new HtmlConverter(mainPart, bookmarkables);
            converter.HtmlStyles.DefaultStyle = converter.HtmlStyles.GetStyle("sc-BodyText");
            Debug.Assert(converter != null, "converter is null");
            Debug.Assert(html != null, "html is null");
            var elements = converter.Parse(html);
            for (int i = 0; i < elements.Count; i++)
            {
                body.Append(elements[i]);
            }
        }

        public void addPara(string style, string content, Boolean indexable)
        {
            Paragraph para = new Paragraph();
            ParagraphProperties paraProps = new ParagraphProperties();
            ParagraphStyleId styleId = new ParagraphStyleId() { Val = style };
            paraProps.Append(styleId);

            Run run = new Run();
            Text text = new Text();
            string escText = System.Web.HttpUtility.HtmlDecode(content);
            text.Text = escText;
            run.Append(text);

            para.Append(paraProps);
            para.Append(run);

            if (indexable == true)
            {
                Run run2 = new Run();
                FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };
                run2.Append(fieldChar1);

                Run run3 = new Run();
                FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode1.Text = " XE \"" + content + "\" ";
                run3.Append(fieldCode1);

                Run run4 = new Run();
                FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.End };
                run4.Append(fieldChar2);

                para.Append(run2);
                para.Append(run3);
                para.Append(run4);
            }

            body.Append(para);
        }

        public void addPara(string style, string content, Boolean indexable, Sitecore.Data.ID id)
        {
            Paragraph para = new Paragraph();
            ParagraphProperties paraProps = new ParagraphProperties();
            ParagraphStyleId styleId = new ParagraphStyleId() { Val = style };
            paraProps.Append(styleId);

            Run run = new Run();
            Text text = new Text();
            string escText = System.Web.HttpUtility.HtmlDecode(content);
            text.Text = escText;
            run.Append(text);

            BookmarkStart bookmarkStart1 = new BookmarkStart() { Name = id.ToString().Replace("-", "").Replace("{", "").Replace("}", ""), Id = bmIdCounter.ToString() };
            BookmarkEnd bookmarkEnd1 = new BookmarkEnd() { Id = bmIdCounter.ToString() };
            bmIdCounter += 1;

            para.Append(paraProps);
            para.Append(bookmarkStart1);
            para.Append(run);
            para.Append(bookmarkEnd1);

            if (indexable == true)
            {
                Run run5 = new Run();
                FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.Begin };
                run5.Append(fieldChar3);

                Run run6 = new Run();
                FieldCode fieldCode4 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode4.Text = " XE \"" + content + "\" ";
                run6.Append(fieldCode4);

                Run run7 = new Run();
                FieldChar fieldChar5 = new FieldChar() { FieldCharType = FieldCharValues.End };
                run7.Append(fieldChar5);

                para.Append(run5);
                para.Append(run6);
                para.Append(run7);
            }

            body.Append(para);
        }

        public void addIndex()
        {
            Paragraph para = new Paragraph();
            ParagraphProperties paraProps = new ParagraphProperties();
            ParagraphStyleId styleId = new ParagraphStyleId() { Val = "Heading2" };
            paraProps.Append(styleId);

            Run run = new Run();
            Text text = new Text();
            text.Text = "Index";
            run.Append(text);

            para.Append(paraProps);
            para.Append(run);

            Run run2 = new Run();
            FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };
            run2.Append(fieldChar1);

            Run run3 = new Run();
            FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
            fieldCode1.Text = " INDEX \\e \"\t\" \\c \"2\" \\z \"1033\" ";
            run3.Append(fieldCode1);

            Run run4 = new Run();
            FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.End };
            run4.Append(fieldChar2);

            para.Append(run2);
            para.Append(run3);
            para.Append(run4);

            insertSection(defaultNumberOfColumns, SectionMarkValues.Continuous);
            body.Append(para);
            insertSection(defaultNumberOfColumns, SectionMarkValues.NextPage);
        }

        public void addToc()
        {
            Paragraph para = new Paragraph();
            ParagraphProperties paraProps = new ParagraphProperties();
            ParagraphStyleId styleId = new ParagraphStyleId() { Val = "TOCTitle" };
            paraProps.Append(styleId);

            Run run = new Run();
            Text text = new Text();
            text.Text = "Table of Contents";
            run.Append(text);

            para.Append(paraProps);
            para.Append(run);

            Run run2 = new Run();
            FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };
            run2.Append(fieldChar1);

            Run run3 = new Run();
            FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
            fieldCode1.Text = @" TOC \o \"1-1\"";
            run3.Append(fieldCode1);

            Run run4 = new Run();
            FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.End };
            run4.Append(fieldChar2);

            para.Append(run2);
            para.Append(run3);
            para.Append(run4);

            body.Append(para);
            insertSection(1, SectionMarkValues.Continuous);
        }

        public void addCatalog(Sitecore.Data.Items.Item item)
        {
            insertSection(1, SectionMarkValues.NextPage);
            addPara("CatalogHeading", item["Title"], true, item.ID);
            string contentSource = item["Content"]; addHTML(contentSource);
            insertSection(1, SectionMarkValues.NextPage);
        }

        public void addContent(Sitecore.Data.Items.Item item)
        {
            breakBefore(item);
            addPara("Heading" + level.ToString(), item.Fields["Title"].Value, true, item.ID);
            string contentSource = item["Content"]; addHTML(contentSource);
            string contentBottomSource = item["Content Bottom"]; addHTML(contentBottomSource);
            breakAfter(item);
        }

        public void addDivision(Sitecore.Data.Items.Item item)
        {
            breakBefore(item);
            addPara("Heading" + level.ToString(), item.Fields["Title"].Value, true, item.ID);

            string contentSource = item.Fields["Content"].Value;
            addHTML(contentSource);

            if (this.institutionItem != null && this.institutionItem.Name.Contains("College-of-Southern-Idaho"))
            {
                RenderAllDivisionFields(item);
            }

            breakAfter(item);
        }

        private void RenderAllDivisionFields(Sitecore.Data.Items.Item item)
        {
            string progMang1 = item["Program Manager 1 Name"]; string progMang1Phone = item["Program Manager 1 Phone"]; string progMang1Email = item["Program Manager 1 Email"];
            string progMang2 = item["Program Manager 2 Name"]; string progMang2Phone = item["Program Manager 2 Phone"]; string progMang2Email = item["Program Manager 2 Email"];
            string focusAdvisor = item["Focus Area Advisor Name"]; string focusAdvisorPhone = item["Focus Area Advisor Phone"]; string focusAdvisorEmail = item["Focus Area Advisor email"];
            string overview = item["Program Overview"]; string website = item["Program Website"]; string employment = item["Gainful Employment"]; string outcomes = item["Program Outcomes"]; string courseReqs = item["Course Performance Requirements"]; string accreditation = item["Program Accreditation Name"];
            string street = item["Program Accreditation Address Street"]; string city = item["Program Accreditation Address City"]; string state = item["Program Accreditation Address State"]; string zip = item["Program Accreditation Address Zip"]; string progPhone = item["Program Accreditation Phone"]; string progFax = item["Program Accreditation Fax"]; string progStatus = item["Program Accreditation Status Description"];
            string application = item["Program Application Required"]; string progReq = item["Program Requirements for Admission"]; string progURL = item["Program Application URL"]; string exitReq = item["Exit Requirements"]; string transfer = item["Transfer Options"]; string sample = item["Sample Career Opportunities"];

            if (!string.IsNullOrWhiteSpace(progMang1)) { addPara("sc-BodyText", "Program Manager: " + progMang1, false); }
            if (!string.IsNullOrWhiteSpace(progMang1Phone)) { addPara("sc-BodyText", "Phone: " + progMang1Phone, false); }
            if (!string.IsNullOrWhiteSpace(progMang1Email)) { addPara("sc-BodyText", "Email: " + progMang1Email, false); }

            if (!string.IsNullOrWhiteSpace(progMang2)) { addPara("sc-BodyText", "Program Manager: " + progMang2, false); }
            if (!string.IsNullOrWhiteSpace(progMang2Phone)) { addPara("sc-BodyText", "Phone: " + progMang2Phone, false); }
            if (!string.IsNullOrWhiteSpace(progMang2Email)) { addPara("sc-BodyText", "Email: " + progMang2Email, false); }

            if (!string.IsNullOrWhiteSpace(focusAdvisor)) { addPara("sc-BodyText", "Focus Area Advisor: " + focusAdvisor, false); }
            if (!string.IsNullOrWhiteSpace(focusAdvisorPhone)) { addPara("sc-BodyText", "Phone: " + focusAdvisorPhone, false); }
            if (!string.IsNullOrWhiteSpace(focusAdvisorEmail)) { addPara("sc-BodyText", "Email: " + focusAdvisorEmail, false); }

            if (!string.IsNullOrWhiteSpace(overview)) { addHTML("Program Overview: " + overview); }
            if (!string.IsNullOrWhiteSpace(website)) { addPara("sc-BodyText", "Visit Program Website: " + website, false); }
            if (!string.IsNullOrWhiteSpace(employment)) { addPara("sc-BodyText", "Gainful Employment: " + employment, false); }
            if (!string.IsNullOrWhiteSpace(outcomes)) { addHTML("Program Outcomes: " + outcomes); }
            if (!string.IsNullOrWhiteSpace(courseReqs)) { addHTML("Course Performance Requirements: " + courseReqs); }
            if (!string.IsNullOrWhiteSpace(accreditation)) { addPara("sc-BodyText", "Program Accreditation: " + accreditation, false); }
            if (!string.IsNullOrWhiteSpace(street)) { addPara("sc-BodyText", "Street: " + street, false); }
            if (!string.IsNullOrWhiteSpace(city)) { addPara("sc-BodyText", "City: " + city, false); }
            if (!string.IsNullOrWhiteSpace(state)) { addPara("sc-BodyText", "State: " + state, false); }
            if (!string.IsNullOrWhiteSpace(zip)) { addPara("sc-BodyText", "Zip: " + zip, false); }
            if (!string.IsNullOrWhiteSpace(progPhone)) { addPara("sc-BodyText", "Phone: " + progPhone, false); }
            if (!string.IsNullOrWhiteSpace(progFax)) { addPara("sc-BodyText", "Fax: " + progFax, false); }
            if (!string.IsNullOrWhiteSpace(progStatus)) { addHTML("Program Accreditation Status/Description: " + progStatus); }
            if (!string.IsNullOrWhiteSpace(application))
            {
                if (application == "{4250D5A4-3A5E-47E1-8CA9-A8AFB0621F4F}") { addPara("sc-BodyText", "Program Application Required: Yes", false); }
                else if (application == "{0F85BA17-9235-4E3F-856C-F37CD40D4B9E}") { addPara("sc-BodyText", "Program Application Required: No", false); }
            }
            if (!string.IsNullOrWhiteSpace(progReq)) { addHTML("Program Requirements for Admission: " + progReq); }
            if (!string.IsNullOrWhiteSpace(progURL)) { addPara("sc-BodyText", progURL, false); }
            if (!string.IsNullOrWhiteSpace(exitReq)) { addHTML("Exit Requirements: " + exitReq); }
            if (!string.IsNullOrWhiteSpace(transfer)) { addHTML("Transfer Options: " + transfer); }
            if (!string.IsNullOrWhiteSpace(sample)) { addHTML("Sample Career Opportunities: " + sample); }
        }

        public void addCoursesFolder(Sitecore.Data.Items.Item item)
        {
            int i = 0;
            string title = item.Fields["Title"].Value;
            if (!int.TryParse(title, out i) && title != "Narrative Courses")
            {
                addPara("Heading" + level.ToString(), item.Fields["Title"].Value, true, item.ID);
                addHTML(item.Fields["Content"].Value);
            }
        }

        public void addCourse(Sitecore.Data.Items.Item item)
        {
            if (hasConfigFile)
            {
                addCourseConfiged(item, database);
            }
            else
            {
                addCourseDefault(item);
            }
        }

        public void addCourseConfiged(Sitecore.Data.Items.Item item, Database database)
        {
            List<courseParagraph> courseParagraphs = new List<courseParagraph>();
            XElement templateNode = (from template in configTree.Descendants("template")
                                     where template.Attribute("name").Value == "Course"
                                     select template).FirstOrDefault();
            List<XElement> paragraphs = (from paragraph in templateNode.Elements()
                                         select paragraph).ToList();

            int i = 0;
            foreach (XElement configPara in paragraphs)
            {
                if (configPara.Name == "paragraph")
                {
                    courseParagraph outputPara = new courseParagraph();

                    string newStyle = "";
                    if (configPara.Attribute("style") != null)
                    {
                        newStyle = configPara.Attribute("style").Value;
                    }
                    outputPara.style = newStyle;

                    if (configPara.Attribute("requires") != null)
                    {
                        outputPara.requires = configPara.Attribute("requires").Value;
                    }
                    else
                    {
                        outputPara.requires = "";
                    }

                    List<XElement> parts = new List<XElement>();
                    parts = (from part in configPara.Descendants()
                             select part).ToList();

                    foreach (XElement part in parts)
                    {
                        paragraphPart newPart = new paragraphPart();
                        if (part.Name != null)
                        {
                            newPart.name = part.Name.ToString();
                        }
                        else
                        {
                            newPart.name = null;
                        }

                        if (part.Attribute("style") != null)
                        {
                            newPart.style = part.Attribute("style").Value;
                        }

                        if (part.Attribute("requires") != null)
                        {
                            newPart.requires = part.Attribute("requires").Value;
                        }

                        if (newPart.name == "field")
                        {
                            string fieldString = part.Attribute("name").Value;
                            if (Replacements.IsIDFixed(fieldString))
                            {
                                Sitecore.Data.ID fieldStringID = Sitecore.Data.ID.Parse(fieldString);
                                newPart.text = item.Fields[fieldString].Value;
                            }
                            else
                            {
                                newPart.text = item.Fields[fieldString].Value;
                            }
                        }

                        if (newPart.name == "text")
                        {
                            string reqs = "";
                            if (part.Attribute("requires") != null)
                            {
                                reqs = part.Attribute("requires").Value;
                            }

                            bool reqsMet = false;

                            if (Replacements.IsIDFixed(reqs))
                            {
                                Sitecore.Data.ID reqId = Sitecore.Data.ID.Parse(reqs);
                                if (item.Fields[reqId].Value != "")
                                {
                                    reqsMet = true;
                                }
                            }
                            else
                            {
                                if (reqs != "")
                                {
                                    if (item.Fields[reqs] != null)
                                    {
                                        if (item.Fields[reqs].Value != "")
                                        {
                                            reqsMet = true;
                                        }
                                    }
                                }
                            }

                            if ((newPart.requires == null) || ((newPart.requires != null) && (reqsMet == true)))
                            {
                                newPart.text = part.Attribute("content").Value;
                                if (part.Attribute("style") != null)
                                {
                                    newPart.style = part.Attribute("style").Value;
                                }
                            }
                        }

                        if (newPart.name == "link")
                        {
                            if (Replacements.IsIDFixed(part.Attribute("name").Value))
                            {
                                string fieldNameOrId = part.Attribute("name").Value;
                                Sitecore.Data.ID fieldID = Sitecore.Data.ID.Parse(fieldNameOrId);

                                if (item.Fields[fieldID] != null)
                                {
                                    string fieldValue = item.Fields[fieldID].Value;
                                    if (fieldValue != "")
                                    {
                                        List<String> linkParts = new List<String>();
                                        Sitecore.Data.ID[] IDArray = Sitecore.Data.ID.ParseArray(fieldValue, true);

                                        foreach (Sitecore.Data.ID id in IDArray)
                                        {
                                            Sitecore.Data.Items.Item targetItem = database.GetItem(id);

                                            if (part.Attribute("field") != null)
                                            {
                                                string targetField = part.Attribute("field").Value;
                                                linkParts.Add(targetItem.Fields[targetField].Value);
                                            }
                                            else
                                            {
                                                linkParts.Add(targetItem.Name);
                                            }
                                        }
                                        string separator = ", ";
                                        if (part.Attribute("separator") != null)
                                        {
                                            separator = part.Attribute("separator").Value;
                                        }
                                        string result = String.Join(separator, linkParts.ToArray());
                                        newPart.text = result;
                                    }
                                }
                            }
                            else
                            {
                                string fieldNameOrId = part.Attribute("name").Value;
                                string fieldValue = item.Fields[fieldNameOrId].Value;
                                List<String> linkParts = new List<String>();
                                Sitecore.Data.ID[] IDArray = Sitecore.Data.ID.ParseArray(fieldValue, true);
                                if (IDArray != null)
                                {
                                    foreach (Sitecore.Data.ID id in IDArray)
                                    {
                                        Sitecore.Data.Items.Item targetItem = database.GetItem(id);

                                        if (part.Attribute("field") != null)
                                        {
                                            string targetField = part.Attribute("field").Value;
                                            linkParts.Add(targetItem.Fields[targetField].Value);
                                        }
                                        else
                                        {
                                            linkParts.Add(targetItem.Name);
                                        }
                                    }
                                    string separator = ", ";
                                    if (part.Attribute("separator") != null)
                                    {
                                        separator = part.Attribute("separator").Value;
                                    }
                                    string result = String.Join(separator, linkParts.ToArray());
                                    newPart.text = result;
                                }
                            }

                            if (part.Attribute("style") != null)
                            {
                                newPart.style = part.Attribute("style").Value;
                            }
                        }

                        outputPara.parts.Add(newPart);
                    }

                    if ((outputPara.requires != null) && (outputPara.requires != ""))
                    {
                        outputPara.reqsMet = false;
                        if (Replacements.IsIDFixed(outputPara.requires) && (item.Fields[Sitecore.Data.ID.Parse(outputPara.requires)] != null) && (item.Fields[Sitecore.Data.ID.Parse(outputPara.requires)].Value != ""))
                        {
                            outputPara.reqsMet = true;
                        }
                        else if ((item.Fields[outputPara.requires] != null) && (item.Fields[outputPara.requires].Value != ""))
                        {
                            outputPara.reqsMet = true;
                        }
                    }
                    else
                    {
                        outputPara.reqsMet = true;
                    }

                    if (outputPara.reqsMet)
                    {
                        Paragraph para = new Paragraph();

                        ParagraphProperties paraProps = new ParagraphProperties();
                        ParagraphStyleId styleId = new ParagraphStyleId() { Val = outputPara.style };
                        paraProps.Append(styleId);
                        para.Append(paraProps);

                        if (i == 0)
                        {
                            BookmarkStart bookmarkStart1 = new BookmarkStart() { Name = item.ID.ToString().Replace("-", "").Replace("{", "").Replace("}", ""), Id = bmIdCounter.ToString() };
                            BookmarkEnd bookmarkEnd1 = new BookmarkEnd() { Id = bmIdCounter.ToString() };
                            bmIdCounter += 1;

                            para.Append(bookmarkStart1);
                            para.Append(bookmarkEnd1);
                            i += 1;
                        }

                        foreach (paragraphPart paraPart in outputPara.parts)
                        {
                            Run run = new Run();
                            Text text = new Text() { Space = SpaceProcessingModeValues.Preserve };
                            string escText = System.Web.HttpUtility.HtmlDecode(paraPart.text);
                            text.Text = escText;
                            if (paraPart.style != null)
                            {
                                RunProperties runProperties = new RunProperties();
                                RunStyle runStyle = new RunStyle() { Val = paraPart.style };
                                runProperties.Append(runStyle);
                                run.Append(runProperties);
                            }
                            run.Append(text);
                            para.Append(run);
                        }
                        body.Append(para);
                    }
                }

                if (configPara.Name == "html")
                {
                    if (Replacements.IsIDFixed(configPara.Attribute("name").Value))
                    {
                        Sitecore.Data.ID htmlID = Sitecore.Data.ID.Parse(configPara.Attribute("name").Value);
                        if (item.Fields[htmlID].Value != "")
                        {
                            string htmlSource = item.Fields[htmlID].Value;
                            addHTML(htmlSource);
                        }
                    }
                    else
                    {
                        if (item.Fields[configPara.Attribute("name").Value].Value != "")
                        {
                            addHTML(item.Fields[configPara.Attribute("name").Value].Value);
                        }
                    }
                }
            }
            foreach (Item child in item.GetChildren())
            {
                if (child.TemplateName == "Outcomes Folder")
                {
                    if (child.Fields["Include in Print"] != null && child["Include in Print"] == "1")
                    {
                        addOutcomes(child);
                    }
                }
            }
        }

        public void addCourseDefault(Sitecore.Data.Items.Item item)
        {
            string headingText = item.Fields["Subject Code"].Value + " " + item.Fields["Course Number"].Value + " - " +
                item.Fields["Course Name"].Value + " (" + item.Fields["Credit Hours Narrative"].Value + ")";

            if ((item.Fields["Course Indicators"] != null) && (item.Fields["Course Indicators"].Value != ""))
            {
                MultilistField indicatorField = item.Fields["Course Indicators"];
                Item[] indicatorItems = indicatorField.GetItems();
                foreach (Item indicator in indicatorItems)
                {
                    if (indicator["Display in Course Descriptions"] == "1")
                    {
                        headingText += indicator["Text"];
                    }
                }
            }
            headingText = headingText.Replace(" ()", "");
            addPara("sc-CourseTitle", headingText, false, item.ID);

            if (!string.IsNullOrWhiteSpace(item.Fields["Instructor"].Value))
            {
                addPara("sc-BodyText", item.Fields["Instructor"].Value, false);
            }

            if (!string.IsNullOrWhiteSpace(item.Fields["Course Description"].Value))
            {
                string descSource = item.Fields["Course Description"].Value;
                addHTML(descSource);
            }

            string preReq = "";
            string coReq = "";
            string crossList = "";
            string distribution = "";
            string offered = "";

            if (item.Fields["Prerequisite Narrative"] != null && !string.IsNullOrWhiteSpace(item.Fields["Prerequisite Narrative"].Value))
            {
                preReq = "Prerequisite: " + item.Fields["Prerequisite Narrative"].Value + ". ";
                preReq = preReq.Replace("..", ".");
            }

            if (item.Fields["Corequisite Narrative"] != null && !string.IsNullOrWhiteSpace(item.Fields["Corequisite Narrative"].Value))
            {
                coReq = "Corequisite: " + item.Fields["Corequisite Narrative"].Value + ". ";
                coReq = coReq.Replace("..", ".");
            }

            if (item.Fields["Cross Listed Courses Narrative"] != null && !string.IsNullOrWhiteSpace(item.Fields["Cross Listed Courses Narrative"].Value))
            {
                crossList = "Crosslisted as: " + item.Fields["Cross Listed Courses Narrative"].Value + ". ";
                crossList = crossList.Replace("..", ".");
            }

            Sitecore.Data.ID offeredId = Sitecore.Data.ID.Parse("{7BA11857-9AF6-4796-AF2A-A5AEC9B98247}");
            if (item.Fields[offeredId] != null && !string.IsNullOrWhiteSpace(item.Fields[offeredId].Value))
            {
                offered = "Offered: " + item.Fields[offeredId].Value + ". ";
                offered = offered.Replace("..", ".");
            }

            Sitecore.Data.ID distributionId = Sitecore.Data.ID.Parse("{49EA1440-FE5A-4682-914D-2C2BFBAC24F2}");
            if (item.Fields[distributionId] != null && !string.IsNullOrWhiteSpace(item.Fields[distributionId].Value))
            {
                distribution = "Distribution: " + item.Fields[distributionId].Value + ". ";
                distribution = distribution.Replace("..", ".");
            }

            string infoString = distribution + preReq + coReq + crossList + offered;
            if (!string.IsNullOrWhiteSpace(infoString))
            {
                addPara("sc-BodyText", infoString, false);
            }

            foreach (Item child in item.GetChildren())
            {
                if (child.TemplateName == "Outcomes Folder")
                {
                    if (child["Include in Print"] == "1" && child.Fields["Include in Print"] != null)
                    {
                        addOutcomes(child);
                    }
                }
            }
        }

        public void addAward(Sitecore.Data.Items.Item item)
        {
            programInserter.AddProgram(item);
        }

        public void addReqs(Sitecore.Data.Items.Item item)
        {
            addPara("sc-RequirementsHeading", item["Title"], false, item.ID);
            addHTML(item["Content"]);

            if (item.HasChildren)
            {
                foreach (Item child in item.GetChildren())
                {
                    if (child["Include in Print"] == "1" || child.Fields["Include in Print"] == null)
                    {
                        if (child.TemplateName == "Requirements List")
                        {
                            addReqList(child, database);
                        }
                        if (child.TemplateName == "Requirements List with Courses")
                        {
                            addReqListWithCourses(child, database);
                        }
                        if (child.TemplateName == "Degree Requirements")
                        {
                            addReqs(child);
                        }
                    }
                }
            }

            addHTML(item["Content Bottom"]);
            addCredits(item, "sc-RequirementsTotal", "Subtotal: ");
        }

        public void addCMFolders(Sitecore.Data.Items.Item item)
        {
            addPara("sc-RequirementsHeading", item["Title"], false, item.ID);
            addHTML(item["Content"]);
            addHTML(item["Content Bottom"]);
        }

        public Table addReqListTable(Item item)
        {
            Table table1 = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TableStyle tableStyle1 = new TableStyle() { Val = "TableGrid" };
            TableWidth tableWidth1 = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };
            tableProperties1.Append(tableStyle1);
            tableProperties1.Append(tableWidth1);

            TableGrid tableGrid1 = new TableGrid();

            //Heinous workaround to include some custom content for a specific institution.
            if (this.institutionItem != null && this.institutionItem.Name.Contains("Rhode-Island-College"))
            {
                GridColumn col1 = new GridColumn() { Width = col1Width };
                GridColumn col2 = new GridColumn() { Width = col2Width };
                GridColumn col3 = new GridColumn() { Width = col3Width };
                GridColumn col4 = new GridColumn() { Width = col4Width };

                tableGrid1.Append(col1);
                tableGrid1.Append(col2);
                tableGrid1.Append(col3);
                tableGrid1.Append(col4);

                table1.Append(tableProperties1);
                table1.Append(tableGrid1);
                return table1;
            }

            if (includeCredits == "1")
            {
                GridColumn col1 = new GridColumn() { Width = col1Width };
                GridColumn col2 = new GridColumn() { Width = col2Width };
                GridColumn col3 = new GridColumn() { Width = col3Width };

                tableGrid1.Append(col1);
                tableGrid1.Append(col2);
                tableGrid1.Append(col3);
            }
            else
            {
                GridColumn col1 = new GridColumn() { Width = col1Width };
                GridColumn col2 = new GridColumn() { Width = col2Width };

                tableGrid1.Append(col1);
                tableGrid1.Append(col2);
            }

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            return table1;
        }

        public void addReqList(Sitecore.Data.Items.Item item, Database database)
        {
            addPara("sc-RequirementsSubheading", item.Fields["Title"].Value, false, item.ID);
            addHTML(item.Fields["Requirement Narrative"].Value);

            if (item.Fields["Course List"].Value != "")
            {
                Table table1 = addReqListTable(item);

                String[] idlist = item.Fields["Course List"].Value.Split('|');
                if ((idlist != null) && (idlist.Length != 0))
                {
                    int i = 0;//for logging only
                    foreach (string reqId in idlist)
                    {
                        i += 1;//for logging only

                        Sitecore.Data.Items.Item reqitem = null;
                        try
                        {
                            reqitem = database.GetItem(Sitecore.Data.ID.Parse(reqId));
                        }
                        catch (Exception ex)
                        {
                            Log.Info("PRINT: Couldn't add requirement list item for " + item.Paths.FullPath, this);
                            Log.Info(ex.Message, this);
                            Log.Info(ex.StackTrace, this);
                            errorNames.Add(item.Paths.FullPath + " (couldn't add a requirement list item)");
                            continue;
                        }

                        if (reqitem == null)
                        {
                            errorNames.Add(item.Paths.FullPath + " (couldn't add a requirement list item)");
                            continue;
                        }

                        string code = reqitem.Fields["Subject Code"].Value;
                        string number = reqitem.Fields["Course Number"].Value;
                        string separator = " ";
                        string scCn = "";

                        try
                        {
                            separator = reqitem.Axes.GetAncestors().Where(x => x.TemplateName == "Courses Folder").First().Fields["Course Name Separator"].Value;
                        }
                        catch
                        {
                            Log.Info("EXPORT DOCX: subject code/course number separator not found for " + reqitem.Paths.FullPath, this);
                        }
                        if (((reqitem.Fields["Subject Code"] != null && reqitem.Fields["Course Number"] != null)) && ((reqitem["Subject Code"] != "") && (reqitem["Course Number"] != "")))
                        {
                            scCn = code + separator + number;
                        }
                        else
                        {
                            scCn = code + number;
                        }

                        if (reqitem.Fields["Cross Listed Courses"].Value != "")
                        {
                            Sitecore.Data.ID[] crossListIds = Sitecore.Data.ID.ParseArray(reqitem.Fields["Cross Listed Courses"].Value);
                            foreach (Sitecore.Data.ID crossListId in crossListIds)
                            {
                                Item crossListItem = database.GetItem(crossListId);
                                scCn += "/" + crossListItem.Fields["Subject Code"].Value + " " + crossListItem.Fields["Course Number"].Value;
                            }
                        }

                        string name = reqitem.Fields["Course Name"].Value;
                        try
                        {
                            if ((reqitem.Fields["Course Indicators"] != null) && (reqitem.Fields["Course Indicators"].Value != ""))
                            {
                                MultilistField indicatorField = reqitem.Fields["Course Indicators"];
                                Item[] indicatorItems = indicatorField.GetItems();
                                foreach (Item indicator in indicatorItems)
                                {
                                    if (indicator["Display in Program Tables"] == "1")
                                    {
                                        name += indicator["Text"];
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Log.Info("PRINT: Couldn't add program indicator for " + name, this);
                            Log.Info(ex.Message, this);
                            Log.Info(ex.StackTrace, this);
                        }

                        bool isRIC = institutionItem != null && institutionItem.Name.Contains("Rhode-Island-College");

                        if ((includeCredits == "1") && (!isRIC))
                        {
                            CreditHours reqch = new CreditHours(reqitem);
                            string creditsString = NormalizeCredits(reqch.NarrativeCreditHours);

                            TableCell cell1 = new TableCell();
                            TableCellProperties tableCellProperties1 = new TableCellProperties();
                            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = col1Width, Type = TableWidthUnitValues.Dxa };
                            tableCellProperties1.Append(tableCellWidth1);
                            Paragraph paragraph1 = new Paragraph();
                            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "sc-Requirement" };
                            paragraphProperties1.Append(paragraphStyleId1);
                            Run run1 = new Run();
                            Text text1 = new Text();
                            text1.Text = scCn;
                            run1.Append(text1);
                            paragraph1.Append(paragraphProperties1);
                            paragraph1.Append(run1);
                            cell1.Append(tableCellProperties1);
                            cell1.Append(paragraph1);

                            TableCell cell2 = new TableCell();
                            TableCellProperties tableCellProperties2 = new TableCellProperties();
                            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = col2Width, Type = TableWidthUnitValues.Dxa };
                            tableCellProperties2.Append(tableCellWidth2);
                            Paragraph paragraph2 = new Paragraph();
                            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "sc-Requirement" };
                            paragraphProperties2.Append(paragraphStyleId2);
                            Run run2 = new Run();
                            Text text2 = new Text();
                            text2.Text = name;
                            run2.Append(text2);
                            paragraph2.Append(paragraphProperties2);
                            paragraph2.Append(run2);
                            cell2.Append(tableCellProperties2);
                            cell2.Append(paragraph2);

                            TableCell cell3 = new TableCell();
                            TableCellProperties tableCellProperties3 = new TableCellProperties();
                            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = col3Width, Type = TableWidthUnitValues.Dxa };
                            tableCellProperties3.Append(tableCellWidth3);
                            Paragraph paragraph3 = new Paragraph();
                            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                            ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() { Val = "sc-RequirementRight" };
                            paragraphProperties3.Append(paragraphStyleId3);
                            Run run3 = new Run();
                            Text text3 = new Text();
                            text3.Text = creditsString;
                            run3.Append(text3);
                            paragraph3.Append(paragraphProperties3);
                            paragraph3.Append(run3);
                            cell3.Append(tableCellProperties3);
                            cell3.Append(paragraph3);

                            TableRow row = new TableRow();
                            row.Append(cell1);
                            row.Append(cell2);
                            row.Append(cell3);
                            table1.Append(row);
                        }
                        if (!(includeCredits == "1") && (!isRIC))
                        {
                            TableCell cell1 = new TableCell();
                            TableCellProperties tableCellProperties1 = new TableCellProperties();
                            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = col1Width, Type = TableWidthUnitValues.Dxa };
                            tableCellProperties1.Append(tableCellWidth1);
                            Paragraph paragraph1 = new Paragraph();
                            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "sc-Requirement" };
                            paragraphProperties1.Append(paragraphStyleId1);
                            Run run1 = new Run();
                            Text text1 = new Text();
                            text1.Text = scCn;
                            run1.Append(text1);
                            paragraph1.Append(paragraphProperties1);
                            paragraph1.Append(run1);
                            cell1.Append(tableCellProperties1);
                            cell1.Append(paragraph1);

                            TableCell cell2 = new TableCell();
                            TableCellProperties tableCellProperties2 = new TableCellProperties();
                            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = col2Width, Type = TableWidthUnitValues.Dxa };
                            tableCellProperties2.Append(tableCellWidth2);
                            Paragraph paragraph2 = new Paragraph();
                            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "sc-Requirement" };
                            paragraphProperties2.Append(paragraphStyleId2);
                            Run run2 = new Run();
                            Text text2 = new Text();
                            text2.Text = name;
                            run2.Append(text2);
                            paragraph2.Append(paragraphProperties2);
                            paragraph2.Append(run2);
                            cell2.Append(tableCellProperties2);
                            cell2.Append(paragraph2);

                            TableRow row = new TableRow();
                            row.Append(cell1);
                            row.Append(cell2);
                            table1.Append(row);
                        }
                        if (isRIC)
                        {
                            CreditHours reqch = new CreditHours(reqitem);
                            string creditsString = NormalizeCredits(reqch.NarrativeCreditHours);
                            string offeredString = NormalizeOfferedString(reqitem["Offered"]);

                            TableCell cell1 = new TableCell();
                            TableCellProperties tableCellProperties1 = new TableCellProperties();
                            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = col1Width, Type = TableWidthUnitValues.Dxa };
                            tableCellProperties1.Append(tableCellWidth1);
                            Paragraph paragraph1 = new Paragraph();
                            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "sc-Requirement" };
                            paragraphProperties1.Append(paragraphStyleId1);
                            Run run1 = new Run();
                            Text text1 = new Text();
                            text1.Text = scCn;
                            run1.Append(text1);
                            paragraph1.Append(paragraphProperties1);
                            paragraph1.Append(run1);
                            cell1.Append(tableCellProperties1);
                            cell1.Append(paragraph1);

                            TableCell cell2 = new TableCell();
                            TableCellProperties tableCellProperties2 = new TableCellProperties();
                            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = col2Width, Type = TableWidthUnitValues.Dxa };
                            tableCellProperties2.Append(tableCellWidth2);
                            Paragraph paragraph2 = new Paragraph();
                            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "sc-Requirement" };
                            paragraphProperties2.Append(paragraphStyleId2);
                            Run run2 = new Run();
                            Text text2 = new Text();
                            text2.Text = name;
                            run2.Append(text2);
                            paragraph2.Append(paragraphProperties2);
                            paragraph2.Append(run2);
                            cell2.Append(tableCellProperties2);
                            cell2.Append(paragraph2);

                            TableCell cell3 = new TableCell();
                            TableCellProperties tableCellProperties3 = new TableCellProperties();
                            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = col3Width, Type = TableWidthUnitValues.Dxa };
                            tableCellProperties3.Append(tableCellWidth3);
                            Paragraph paragraph3 = new Paragraph();
                            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                            ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() { Val = "sc-RequirementRight" };
                            paragraphProperties3.Append(paragraphStyleId3);
                            Run run3 = new Run();
                            Text text3 = new Text();
                            text3.Text = creditsString;
                            run3.Append(text3);
                            paragraph3.Append(paragraphProperties3);
                            paragraph3.Append(run3);
                            cell3.Append(tableCellProperties3);
                            cell3.Append(paragraph3);

                            TableCell cell4 = new TableCell();
                            TableCellProperties tableCellProperties4 = new TableCellProperties();
                            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = col4Width, Type = TableWidthUnitValues.Dxa };
                            tableCellProperties4.Append(tableCellWidth4);
                            Paragraph paragraph4 = new Paragraph();
                            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
                            ParagraphStyleId paragraphStyleId4 = new ParagraphStyleId() { Val = "sc-Requirement" };
                            paragraphProperties4.Append(paragraphStyleId4);
                            Run run4 = new Run();
                            Text text4 = new Text();
                            text4.Text = offeredString;
                            run4.Append(text4);
                            paragraph4.Append(paragraphProperties4);
                            paragraph4.Append(run4);
                            cell4.Append(tableCellProperties4);
                            cell4.Append(paragraph4);

                            TableRow row = new TableRow();
                            row.Append(cell1);
                            row.Append(cell2);
                            row.Append(cell3);
                            row.Append(cell4);
                            table1.Append(row);
                        }
                    }
                }
                body.Append(table1);
            }

            addCredits(item, "sc-Subtotal", "Subtotal: ");

            string noteSource = item.Fields["Requirement Note"].Value;
            addHTML(noteSource);

            if (item.HasChildren)
            {
                foreach (Item child in item.GetChildren())
                {
                    addReqList(child, database);
                }
            }
            if ((item.Fields["Milestones"] != null) && (item["Milestones"] != ""))
            {
                addPara("sc-RequirementsSubheading", "Milestones", false);
                addHTML(item.Fields["Milestones"].Value);
            }
        }

        public void addReqListWithCourses(Sitecore.Data.Items.Item item, Database database)
        {
            addPara("sc-RequirementsSubheading", item.Fields["Title"].Value, false, item.ID);
            addHTML(item.Fields["Requirement Narrative"].Value);

            if (item.Fields["Course List"].Value != "")
            {
                Table table1 = addReqListTable(item);

                String[] idlist = item.Fields["Course List"].Value.Split('|');
                if ((idlist != null) && (idlist.Length != 0))
                {
                    int i = 0;//for logging only
                    foreach (string reqId in idlist)
                    {
                        i += 1;//for logging only

                        Sitecore.Data.Items.Item reqitem = null;
                        try
                        {
                            reqitem = database.GetItem(Sitecore.Data.ID.Parse(reqId));
                        }
                        catch (Exception ex)
                        {
                            Log.Info("PRINT: Couldn't add requirement list item for " + item.Paths.FullPath, this);
                            Log.Info(ex.Message, this);
                            Log.Info(ex.StackTrace, this);
                            errorNames.Add(item.Paths.FullPath + " (couldn't add a requirement list item)");
                            continue;
                        }

                        if (reqitem == null)
                        {
                            errorNames.Add(item.Paths.FullPath + " (couldn't add a requirement list item)");
                            continue;
                        }

                        string code = reqitem.Fields["Subject Code"].Value;
                        string number = reqitem.Fields["Course Number"].Value;
                        string separator = " ";
                        string scCn = "";

                        try
                        {
                            separator = reqitem.Axes.GetAncestors().Where(x => x.TemplateName == "Courses Folder").First().Fields["Course Name Separator"].Value;
                        }
                        catch
                        {
                            Log.Info("EXPORT DOCX: subject code/course number separator not found for " + reqitem.Paths.FullPath, this);
                        }
                        if (((reqitem.Fields["Subject Code"] != null && reqitem.Fields["Course Number"] != null)) && ((reqitem["Subject Code"] != "") && (reqitem["Course Number"] != "")))
                        {
                            scCn = code + separator + number;
                        }
                        else
                        {
                            scCn = code + number;
                        }

                        if (reqitem.Fields["Cross Listed Courses"].Value != "")
                        {
                            Sitecore.Data.ID[] crossListIds = Sitecore.Data.ID.ParseArray(reqitem.Fields["Cross Listed Courses"].Value);
                            foreach (Sitecore.Data.ID crossListId in crossListIds)
                            {
                                Item crossListItem = database.GetItem(crossListId);
                                scCn += "/" + crossListItem.Fields["Subject Code"].Value + " " + crossListItem.Fields["Course Number"].Value;
                            }
                        }

                        string name = reqitem.Fields["Course Name"].Value;
                        try
                        {
                            if ((reqitem.Fields["Course Indicators"] != null) && (reqitem.Fields["Course Indicators"].Value != ""))
                            {
                                MultilistField indicatorField = reqitem.Fields["Course Indicators"];
                                Item[] indicatorItems = indicatorField.GetItems();
                                foreach (Item indicator in indicatorItems)
                                {
                                    if (indicator["Display in Program Tables"] == "1")
                                    {
                                        name += indicator["Text"];
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Log.Info("PRINT: Couldn't add program indicator for " + name, this);
                            Log.Info(ex.Message, this);
                            Log.Info(ex.StackTrace, this);
                        }

                        bool isRIC = institutionItem != null && institutionItem.Name.Contains("Rhode-Island-College");

                        if ((includeCredits == "1") && (!isRIC))
                        {
                            CreditHours reqch = new CreditHours(reqitem);
                            string creditsString = NormalizeCredits(reqch.NarrativeCreditHours);

                            TableCell cell1 = new TableCell();
                            TableCellProperties tableCellProperties1 = new TableCellProperties();
                            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = col1Width, Type = TableWidthUnitValues.Dxa };
                            tableCellProperties1.Append(tableCellWidth1);
                            Paragraph paragraph1 = new Paragraph();
                            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "sc-Requirement" };
                            paragraphProperties1.Append(paragraphStyleId1);
                            Run run1 = new Run();
                            Text text1 = new Text();
                            text1.Text = scCn;
                            run1.Append(text1);
                            paragraph1.Append(paragraphProperties1);
                            paragraph1.Append(run1);
                            cell1.Append(tableCellProperties1);
                            cell1.Append(paragraph1);

                            TableCell cell2 = new TableCell();
                            TableCellProperties tableCellProperties2 = new TableCellProperties();
                            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = col2Width, Type = TableWidthUnitValues.Dxa };
                            tableCellProperties2.Append(tableCellWidth2);
                            Paragraph paragraph2 = new Paragraph();
                            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "sc-Requirement" };
                            paragraphProperties2.Append(paragraphStyleId2);
                            Run run2 = new Run();
                            Text text2 = new Text();
                            text2.Text = name;
                            run2.Append(text2);
                            paragraph2.Append(paragraphProperties2);
                            paragraph2.Append(run2);
                            cell2.Append(tableCellProperties2);
                            cell2.Append(paragraph2);

                            TableCell cell3 = new TableCell();
                            TableCellProperties tableCellProperties3 = new TableCellProperties();
                            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = col3Width, Type = TableWidthUnitValues.Dxa };
                            tableCellProperties3.Append(tableCellWidth3);
                            Paragraph paragraph3 = new Paragraph();
                            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                            ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() { Val = "sc-RequirementRight" };
                            paragraphProperties3.Append(paragraphStyleId3);
                            Run run3 = new Run();
                            Text text3 = new Text();
                            text3.Text = creditsString;
                            run3.Append(text3);
                            paragraph3.Append(paragraphProperties3);
                            paragraph3.Append(run3);
                            cell3.Append(tableCellProperties3);
                            cell3.Append(paragraph3);

                            TableRow row = new TableRow();
                            row.Append(cell1);
                            row.Append(cell2);
                            row.Append(cell3);
                            table1.Append(row);
                        }
                        if (!(includeCredits == "1") && (!isRIC))
                        {
                            TableCell cell1 = new TableCell();
                            TableCellProperties tableCellProperties1 = new TableCellProperties();
                            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = col1Width, Type = TableWidthUnitValues.Dxa };
                            tableCellProperties1.Append(tableCellWidth1);
                            Paragraph paragraph1 = new Paragraph();
                            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "sc-Requirement" };
                            paragraphProperties1.Append(paragraphStyleId1);
                            Run run1 = new Run();
                            Text text1 = new Text();
                            text1.Text = scCn;
                            run1.Append(text1);
                            paragraph1.Append(paragraphProperties1);
                            paragraph1.Append(run1);
                            cell1.Append(tableCellProperties1);
                            cell1.Append(paragraph1);

                            TableCell cell2 = new TableCell();
                            TableCellProperties tableCellProperties2 = new TableCellProperties();
                            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = col2Width, Type = TableWidthUnitValues.Dxa };
                            tableCellProperties2.Append(tableCellWidth2);
                            Paragraph paragraph2 = new Paragraph();
                            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "sc-Requirement" };
                            paragraphProperties2.Append(paragraphStyleId2);
                            Run run2 = new Run();
                            Text text2 = new Text();
                            text2.Text = name;
                            run2.Append(text2);
                            paragraph2.Append(paragraphProperties2);
                            paragraph2.Append(run2);
                            cell2.Append(tableCellProperties2);
                            cell2.Append(paragraph2);

                            TableRow row = new TableRow();
                            row.Append(cell1);
                            row.Append(cell2);
                            table1.Append(row);
                        }
                        if (isRIC)
                        {
                            CreditHours reqch = new CreditHours(reqitem);
                            string creditsString = NormalizeCredits(reqch.NarrativeCreditHours);
                            string offeredString = NormalizeOfferedString(reqitem["Offered"]);

                            TableCell cell1 = new TableCell();
                            TableCellProperties tableCellProperties1 = new TableCellProperties();
                            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = col1Width, Type = TableWidthUnitValues.Dxa };
                            tableCellProperties1.Append(tableCellWidth1);
                            Paragraph paragraph1 = new Paragraph();
                            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "sc-Requirement" };
                            paragraphProperties1.Append(paragraphStyleId1);
                            Run run1 = new Run();
                            Text text1 = new Text();
                            text1.Text = scCn;
                            run1.Append(text1);
                            paragraph1.Append(paragraphProperties1);
                            paragraph1.Append(run1);
                            cell1.Append(tableCellProperties1);
                            cell1.Append(paragraph1);

                            TableCell cell2 = new TableCell();
                            TableCellProperties tableCellProperties2 = new TableCellProperties();
                            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = col2Width, Type = TableWidthUnitValues.Dxa };
                            tableCellProperties2.Append(tableCellWidth2);
                            Paragraph paragraph2 = new Paragraph();
                            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "sc-Requirement" };
                            paragraphProperties2.Append(paragraphStyleId2);
                            Run run2 = new Run();
                            Text text2 = new Text();
                            text2.Text = name;
                            run2.Append(text2);
                            paragraph2.Append(paragraphProperties2);
                            paragraph2.Append(run2);
                            cell2.Append(tableCellProperties2);
                            cell2.Append(paragraph2);

                            TableCell cell3 = new TableCell();
                            TableCellProperties tableCellProperties3 = new TableCellProperties();
                            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = col3Width, Type = TableWidthUnitValues.Dxa };
                            tableCellProperties3.Append(tableCellWidth3);
                            Paragraph paragraph3 = new Paragraph();
                            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                            ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() { Val = "sc-RequirementRight" };
                            paragraphProperties3.Append(paragraphStyleId3);
                            Run run3 = new Run();
                            Text text3 = new Text();
                            text3.Text = creditsString;
                            run3.Append(text3);
                            paragraph3.Append(paragraphProperties3);
                            paragraph3.Append(run3);
                            cell3.Append(tableCellProperties3);
                            cell3.Append(paragraph3);

                            TableCell cell4 = new TableCell();
                            TableCellProperties tableCellProperties4 = new TableCellProperties();
                            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = col4Width, Type = TableWidthUnitValues.Dxa };
                            tableCellProperties4.Append(tableCellWidth4);
                            Paragraph paragraph4 = new Paragraph();
                            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
                            ParagraphStyleId paragraphStyleId4 = new ParagraphStyleId() { Val = "sc-Requirement" };
                            paragraphProperties4.Append(paragraphStyleId4);
                            Run run4 = new Run();
                            Text text4 = new Text();
                            text4.Text = offeredString;
                            run4.Append(text4);
                            paragraph4.Append(paragraphProperties4);
                            paragraph4.Append(run4);
                            cell4.Append(tableCellProperties4);
                            cell4.Append(paragraph4);

                            TableRow row = new TableRow();
                            row.Append(cell1);
                            row.Append(cell2);
                            row.Append(cell3);
                            row.Append(cell4);
                            table1.Append(row);
                        }
                    }
                }
                body.Append(table1);
            }

            addCredits(item, "sc-Subtotal", "Subtotal: ");

            string noteSource = item.Fields["Requirement Note"].Value;
            addHTML(noteSource);

            if (item.HasChildren)
            {
                foreach (Item child in item.GetChildren())
                {
                    addReqListWithCourses(child, database);
                }
            }
            if ((item.Fields["Milestones"] != null) && (item["Milestones"] != ""))
            {
                addPara("sc-RequirementsSubheading", "Milestones", false);
                addHTML(item.Fields["Milestones"].Value);
            }
        }

        public void addCredits(Sitecore.Data.Items.Item item, string style, string label)
        {
            CreditHours ch = new CreditHours(item);
            string credits = "";
            if (item["Auto Calculate Credit Hours"] != "1")
            {
                credits = ch.NarrativeCreditHours;
            }
            else
            {
                credits = ch.TotalCreditHours;
                if (credits == "0")
                {
                    credits = "";
                }
            }

            if (!string.IsNullOrEmpty(credits) && (
                (
                    (calcSubtotals == "1") &&
                    (item.TemplateName == "Requirements List" || item.TemplateName == "Requirements List with Courses")
                ) ||
                (
                    (calcDegReq == "1") &&
                    (item.TemplateName == "Degree Requirements" || item.TemplateName == "Degree-Requirements")
                ) ||
                (
                    (calcTotals == "1") &&
                    (item.TemplateName == "Degree" || item.TemplateName == "Certificate" || item.TemplateName == "Minor" || item.TemplateName == "Narrative with Course Table")
                )
            ))
            {
                string creditString = label + credits;
                addPara(style, creditString, false);
            }
        }

        public void addOutcomes(Item item)
        {
            if ((item.TemplateName != "Program Outcomes") && (item.TemplateName != "Program-Outcomes") && (item.TemplateName != "Outcomes Folder") && (item.TemplateName != "Course Outcomes"))
            {
                return;
            }

            string title = item["Title"];
            if (!string.IsNullOrEmpty(title))
            {
                addPara("sc-RequirementsHeading", item["Title"], false, item.ID);
            }

            if (item.HasChildren)
            {
                foreach (Item child in item.GetChildren())
                {
                    if (child.TemplateName == "Program Goals" || child.TemplateName == "Program-Goals")
                    {
                        addPara("sc-OutcomeGoal", child["Goal"], false);
                    }
                    else if (child.TemplateName == "Outcome")
                    {
                        addPara("sc-OutcomeGoal", child["Outcome"], false);
                    }
                }
            }
        }

        public void addOther(Sitecore.Data.Items.Item item)
        {
            string filler = "*** " + item.TemplateName + " ***";
            addPara("red", filler, false);
            string filler2 = "*** " + item.Name + " ***";
            addPara("red", filler2, false);
        }
    }

    public class PrintSection : Sitecore.Shell.Framework.Commands.Command
    {
        public override void Execute(Sitecore.Shell.Framework.Commands.CommandContext context)
        {
            if (context.Items.Length == 1)
            {
                Sitecore.Data.Items.Item item = context.Items[0];
                System.Collections.Specialized.NameValueCollection parameters = new System.Collections.Specialized.NameValueCollection();
                parameters["name"] = item.Name;
                parameters["ID"] = item.ID.ToString();
                parameters["database"] = item.Database?.Name ?? "master";
                parameters["path"] = item.Paths.FullPath;

                UrlString url = new UrlString("/");
                url["sc_itemid"] = "{81BA340F-7CFE-4236-B166-CDDD171A0AD2}"; //docgen
                url["startitem"] = item.ID.ToString();
                url["sc_lang"] = item.Language.CultureInfo.Name;
                SheerResponse.Eval("window.open('" + url.ToString() + "', '_blank')");
            }
        }
    }
}

