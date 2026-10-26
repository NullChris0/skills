using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace DeliveryDoc;

public static class Program
{
    private const int A4W = 11906;
    private const int A4H = 16838;

    private static class C
    {
        public const string Primary = "344B46";
        public const string Accent = "B69563";
        public const string Dark = "24312E";
        public const string Light = "8A9792";
        public const string Border = "D7DDD9";
        public const string Panel = "F3F6F4";
        public const string Header = "DDE7E1";
        public const string White = "FFFFFF";
    }

    private static uint _docPrId = 1;

    public static int Main(string[] args)
    {
        if (args.Length != 2)
        {
            Console.Error.WriteLine(
                "Usage: dotnet run -- <delivery-model.json> <output.docx>");
            return 2;
        }

        try
        {
            string modelPath = Path.GetFullPath(args[0]);
            string outputPath = Path.GetFullPath(args[1]);
            JsonObject root = JsonNode.Parse(File.ReadAllText(modelPath)) as JsonObject
                ?? throw new InvalidDataException("Delivery model must be a JSON object.");
            JsonObject document = root["document"] as JsonObject
                ?? throw new InvalidDataException("Delivery model is missing the document object.");
            JsonArray sections = document["sections"] as JsonArray
                ?? throw new InvalidDataException("document.sections must be an array.");

            string outputDirectory = Path.GetDirectoryName(outputPath)
                ?? throw new InvalidDataException("Output path has no parent directory.");
            Directory.CreateDirectory(outputDirectory);
            ValidateImages(sections, outputDirectory);

            using var package = WordprocessingDocument.Create(
                outputPath,
                WordprocessingDocumentType.Document);
            var main = package.AddMainDocumentPart();
            main.Document = new Document(new Body());
            var body = main.Document.Body!;

            AddStyles(main);
            AddCoreProperties(package, document);
            (string headerId, string footerId) =
                AddHeaderFooter(main, S(document["title"]));

            AddPlainCover(body, document);
            AddToc(body, sections);
            AddBodyContent(body, main, sections, outputDirectory);
            AddFinalSection(body, headerId, footerId);
            SetUpdateFieldsOnOpen(main);

            main.Document.Save();
            package.Save();
            Console.WriteLine($"Generated: {outputPath}");
            return 0;
        }
        catch (Exception error)
        {
            Console.Error.WriteLine($"ERROR: {error.Message}");
            return 1;
        }
    }

    private static string S(JsonNode? value, string fallback = "")
    {
        if (value is null) return fallback;
        try { return value.GetValue<string>(); }
        catch { return value.ToJsonString(); }
    }

    private static int I(JsonNode? value)
    {
        if (value is null) return 0;
        try { return value.GetValue<int>(); }
        catch
        {
            int.TryParse(S(value), out int result);
            return result;
        }
    }

    private static void ValidateImages(JsonArray sections, string outputDirectory)
    {
        foreach (JsonNode? section in sections)
        {
            if (section?["images"] is not JsonArray images) continue;
            foreach (JsonNode? image in images)
            {
                string relativePath = S(image?["path"]);
                string imagePath = ResolveImagePath(outputDirectory, relativePath);
                if (!File.Exists(imagePath))
                    throw new FileNotFoundException(
                        $"Image not found: {relativePath}",
                        imagePath);
            }
        }
    }

    private static string ResolveImagePath(
        string outputDirectory,
        string relativePath)
    {
        if (string.IsNullOrWhiteSpace(relativePath))
            throw new InvalidDataException("Image path must not be empty.");
        if (Path.IsPathRooted(relativePath))
            throw new InvalidDataException(
                $"Image path must be relative to the DOCX directory: {relativePath}");
        return Path.GetFullPath(Path.Combine(outputDirectory, relativePath));
    }

    private static void AddStyles(MainDocumentPart main)
    {
        var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles();
        Styles styles = stylesPart.Styles;

        styles.Append(new Style(
            new StyleName { Val = "Normal" },
            new StyleParagraphProperties(
                new SpacingBetweenLines
                {
                    Before = "0",
                    After = "140",
                    Line = "360",
                    LineRule = LineSpacingRuleValues.Auto
                },
                new Indentation { FirstLine = "420" }),
            new StyleRunProperties(
                new RunFonts
                {
                    Ascii = "Times New Roman",
                    HighAnsi = "Times New Roman",
                    EastAsia = "宋体"
                },
                new Color { Val = "000000" },
                new FontSize { Val = "21" }))
        {
            Type = StyleValues.Paragraph,
            StyleId = "Normal",
            Default = true
        });

        styles.Append(CreateHeadingStyle(
            "Heading1", "heading 1", 0, "32", "0", "180", true));
        styles.Append(CreateHeadingStyle(
            "Heading2", "heading 2", 1, "28", "320", "120", false));
        styles.Append(CreateHeadingStyle(
            "Heading3", "heading 3", 2, "24", "280", "100", false));
        styles.Append(CreateHeadingStyle(
            "Heading4", "heading 4", 3, "22", "240", "80", false));
        styles.Append(CreateHeadingStyle(
            "Heading5", "heading 5", 4, "21", "200", "60", false));

        styles.Append(new Style(
            new StyleName { Val = "Caption" },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new SpacingBetweenLines { Before = "60", After = "300" },
                new Indentation { FirstLine = "0" },
                new Justification { Val = JustificationValues.Center }),
            new StyleRunProperties(
                new RunFonts
                {
                    Ascii = "Times New Roman",
                    HighAnsi = "Times New Roman",
                    EastAsia = "宋体"
                },
                new Color { Val = "000000" },
                new FontSize { Val = "20" }))
        {
            Type = StyleValues.Paragraph,
            StyleId = "Caption"
        });

        styles.Append(CreateTocStyle("TOC1", "toc 1", true, "0", "180"));
        styles.Append(CreateTocStyle("TOC2", "toc 2", false, "360", "70"));
        styles.Append(CreateTocStyle("TOC3", "toc 3", false, "720", "40"));
        styles.Append(CreateTocStyle("TOC4", "toc 4", false, "1080", "20"));
    }

    private static Style CreateHeadingStyle(
        string id,
        string name,
        int level,
        string size,
        string before,
        string after,
        bool pageBreak)
    {
        var properties = new StyleParagraphProperties(
            new KeepNext(),
            new KeepLines());
        if (pageBreak) properties.Append(new PageBreakBefore());
        properties.Append(new SpacingBetweenLines { Before = before, After = after });
        properties.Append(new Indentation { FirstLine = "0" });
        properties.Append(new OutlineLevel { Val = level });
        var runProperties = new StyleRunProperties(new RunFonts
        {
            Ascii = "Times New Roman",
            HighAnsi = "Times New Roman",
            EastAsia = "宋体"
        });
        runProperties.Append(new Bold());
        runProperties.Append(new Color { Val = "000000" });
        runProperties.Append(new FontSize { Val = size });
        return new Style(
            new StyleName { Val = name },
            new BasedOn { Val = "Normal" },
            properties,
            runProperties)
        {
            Type = StyleValues.Paragraph,
            StyleId = id
        };
    }

    private static Style CreateTocStyle(
        string id,
        string name,
        bool bold,
        string indent,
        string before)
    {
        var runProperties = new StyleRunProperties(new RunFonts
        {
            Ascii = "Times New Roman",
            HighAnsi = "Times New Roman",
            EastAsia = "宋体"
        });
        if (bold) runProperties.Append(new Bold());
        runProperties.Append(new Color { Val = "000000" });
        runProperties.Append(new FontSize { Val = "20" });
        return new Style(
            new StyleName { Val = name },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new Tabs(new TabStop
                {
                    Val = TabStopValues.Right,
                    Leader = TabStopLeaderCharValues.Dot,
                    Position = 9000
                }),
                new SpacingBetweenLines { Before = before, After = "30" },
                new Indentation { Left = indent, FirstLine = "0" }),
            runProperties)
        {
            Type = StyleValues.Paragraph,
            StyleId = id
        };
    }

    private static void AddCoreProperties(
        WordprocessingDocument package,
        JsonObject document)
    {
        package.PackageProperties.Title = S(document["title"]);
        package.PackageProperties.Subject = "需求说明书";
    }

    private static (string HeaderId, string FooterId) AddHeaderFooter(
        MainDocumentPart main,
        string title)
    {
        var headerPart = main.AddNewPart<HeaderPart>();
        string headerId = main.GetIdOfPart(headerPart);
        headerPart.Header = new Header(new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { After = "0" },
                new Indentation { FirstLine = "0" },
                new Justification { Val = JustificationValues.Right }),
            new Run(
                new RunProperties(
                    new RunFonts
                    {
                        Ascii = "Calibri",
                        HighAnsi = "Calibri",
                        EastAsia = "宋体"
                    },
                    new Color { Val = C.Light },
                    new FontSize { Val = "16" }),
                new Text(title))));

        var footerPart = main.AddNewPart<FooterPart>();
        string footerId = main.GetIdOfPart(footerPart);
        var paragraph = new Paragraph(new ParagraphProperties(
            new SpacingBetweenLines { Before = "0", After = "0" },
            new Indentation { FirstLine = "0" },
            new Justification { Val = JustificationValues.Center }));
        AppendFooterField(paragraph, " PAGE ", "1");
        footerPart.Footer = new Footer(paragraph);
        return (headerId, footerId);
    }

    private static void AppendFooterField(
        Paragraph paragraph,
        string code,
        string placeholder)
    {
        paragraph.Append(new Run(FooterRunProperties(), new FieldChar
        {
            FieldCharType = FieldCharValues.Begin
        }));
        paragraph.Append(new Run(FooterRunProperties(), new FieldCode(code)
        {
            Space = SpaceProcessingModeValues.Preserve
        }));
        paragraph.Append(new Run(FooterRunProperties(), new FieldChar
        {
            FieldCharType = FieldCharValues.Separate
        }));
        paragraph.Append(new Run(FooterRunProperties(), new Text(placeholder)));
        paragraph.Append(new Run(FooterRunProperties(), new FieldChar
        {
            FieldCharType = FieldCharValues.End
        }));
    }

    private static RunProperties FooterRunProperties()
    {
        return new RunProperties(
            new RunFonts
            {
                Ascii = "Calibri",
                HighAnsi = "Calibri",
                EastAsia = "宋体"
            },
            new Color { Val = "000000" },
            new FontSize { Val = "18" });
    }

    private static void AddPlainCover(Body body, JsonObject document)
    {
        body.Append(new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "6800", After = "120" },
                new Indentation { FirstLine = "0" },
                new Justification { Val = JustificationValues.Center }),
            CoverRun(S(document["title"]), "32", true)));

        body.Append(new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "0", After = "180" },
                new Indentation { FirstLine = "0" },
                new Justification { Val = JustificationValues.Center }),
            CoverRun("需求说明书", "28", true)));

        body.Append(new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "180", After = "0" },
                new Indentation { FirstLine = "0" },
                new Justification { Val = JustificationValues.Center }),
            CoverRun(MonthText(S(document["date"])), "22", false)));

        body.Append(new Paragraph(new ParagraphProperties(
            new SectionProperties(
                new SectionType { Val = SectionMarkValues.NextPage },
                new PageSize
                {
                    Width = (UInt32Value)(uint)A4W,
                    Height = (UInt32Value)(uint)A4H
                },
                new PageMargin
                {
                    Top = 0,
                    Right = 0,
                    Bottom = 0,
                    Left = 0,
                    Header = 0,
                    Footer = 0
                }))));
    }

    private static string MonthText(string value)
    {
        Match match = Regex.Match(value, @"(\d{4})年(\d{1,2})月");
        if (!match.Success)
            throw new InvalidDataException("document.date 必须包含 YYYY年M月");
        return $"{match.Groups[1].Value}年{match.Groups[2].Value}月";
    }

    private static Run CoverRun(string text, string size, bool bold)
    {
        var properties = new RunProperties(new RunFonts
        {
            Ascii = "Times New Roman",
            HighAnsi = "Times New Roman",
            EastAsia = "宋体"
        });
        if (bold) properties.Append(new Bold());
        properties.Append(new Color { Val = "000000" });
        properties.Append(new FontSize { Val = size });
        return new Run(properties, new Text(text));
    }

    private static void AddToc(Body body, JsonArray sections)
    {
        body.Append(new Paragraph(
            new ParagraphProperties(
                new PageBreakBefore(),
                new SpacingBetweenLines { Before = "0", After = "280" },
                new Indentation { FirstLine = "0" }),
            new Run(
                new RunProperties(
                    new RunFonts
                    {
                        Ascii = "Calibri",
                        HighAnsi = "Calibri",
                        EastAsia = "微软雅黑"
                    },
                    new Bold(),
                    new Color { Val = C.Primary },
                    new FontSize { Val = "34" }),
                new Text("目录"))));

        body.Append(new Paragraph(
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode(" TOC \\o \"1-4\" \\h \\z \\u ")
            {
                Space = SpaceProcessingModeValues.Preserve
            }),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate })));

        foreach (JsonNode? section in sections)
        {
            int level = I(section?["level"]);
            if (level is >= 1 and <= 4)
                body.Append(CreateTocEntry(
                    S(section?["heading"]),
                    level,
                    S(section?["page_hint"], "1")));
        }

        body.Append(new Paragraph(new Run(new FieldChar
        {
            FieldCharType = FieldCharValues.End
        })));
    }

    private static Paragraph CreateTocEntry(
        string text,
        int level,
        string pageHint)
    {
        return new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = $"TOC{level}" }),
            new Run(new Text(text)),
            new Run(new TabChar()),
            new Run(new Text(pageHint)));
    }

    private static void AddBodyContent(
        Body body,
        MainDocumentPart main,
        JsonArray sections,
        string outputDirectory)
    {
        for (int index = 0; index < sections.Count; index++)
        {
            if (sections[index] is not JsonObject sectionObject)
                throw new InvalidDataException("Each section must be an object.");

            int level = I(sectionObject["level"]);
            JsonObject? next = index + 1 < sections.Count
                ? sections[index + 1] as JsonObject
                : null;
            bool tight = next is not null && I(next["level"]) > level;

            body.Append(CreateHeading(
                S(sectionObject["heading"]),
                level,
                tight));

            if (sectionObject["paras"] is JsonArray paragraphs)
            {
                foreach (JsonNode? paragraph in paragraphs)
                    body.Append(CreateParagraph(S(paragraph)));
            }

            if (sectionObject["tables"] is JsonArray tables)
            {
                foreach (JsonNode? tableNode in tables)
                {
                    if (tableNode is not JsonObject table)
                        throw new InvalidDataException("Each table must be an object.");
                    string[] headers = (table["headers"] as JsonArray
                        ?? throw new InvalidDataException("Table headers must be an array."))
                        .Select(value => S(value))
                        .ToArray();
                    int[] widths = (table["widths"] as JsonArray
                        ?? throw new InvalidDataException("Table widths must be an array."))
                        .Select(value => I(value))
                        .ToArray();
                    List<string[]> rows = (table["rows"] as JsonArray
                        ?? throw new InvalidDataException("Table rows must be an array."))
                        .Select(row => (row as JsonArray
                            ?? throw new InvalidDataException("Each table row must be an array."))
                            .Select(value => S(value))
                            .ToArray())
                        .ToList();

                    if (headers.Length != widths.Length)
                        throw new InvalidDataException(
                            "Table headers and widths must have the same length.");
                    if (rows.Any(row => row.Length != widths.Length))
                        throw new InvalidDataException(
                            "Each table row must match the header width count.");

                    body.Append(CreateBorderedTable(headers, widths, rows));
                    body.Append(new Paragraph(new ParagraphProperties(
                        new SpacingBetweenLines { After = "100" })));
                }
            }

            if (sectionObject["meta"] is JsonArray metadata && metadata.Count > 0)
            {
                Table table = CreateTableWithWidths(new[] { 1900, 6400 });
                foreach (JsonNode? item in metadata)
                {
                    if (item is JsonArray pair && pair.Count >= 2)
                        AddInfoRow(table, S(pair[0]), S(pair[1]));
                }
                body.Append(table);
                body.Append(new Paragraph(new ParagraphProperties(
                    new SpacingBetweenLines { After = "100" })));
            }

            if (sectionObject["images"] is JsonArray images)
            {
                foreach (JsonNode? image in images)
                {
                    string path = ResolveImagePath(
                        outputDirectory,
                        S(image?["path"]));
                    string caption = S(image?["caption"]);
                    AddInlineImage(
                        body,
                        main,
                        path,
                        string.IsNullOrWhiteSpace(caption)
                            ? Path.GetFileName(path)
                            : caption);
                    if (!string.IsNullOrWhiteSpace(caption))
                        body.Append(CreateCaption(caption));
                }
            }
        }
    }

    private static Paragraph CreateHeading(string text, int level, bool tight)
    {
        string style = level switch
        {
            <= 1 => "Heading1",
            2 => "Heading2",
            3 => "Heading3",
            4 => "Heading4",
            _ => "Heading5"
        };
        var properties = new ParagraphProperties(
            new ParagraphStyleId { Val = style });
        if (tight)
        {
            string before = level switch
            {
                <= 1 => "0",
                2 => "320",
                3 => "280",
                4 => "240",
                _ => "200"
            };
            properties.Append(new SpacingBetweenLines
            {
                Before = before,
                After = "0"
            });
        }
        return new Paragraph(
            properties,
            new Run(new Text(text)));
    }

    private static Paragraph CreateParagraph(string text)
    {
        return new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines
                {
                    After = "140",
                    Line = "360",
                    LineRule = LineSpacingRuleValues.Auto
                },
                new Indentation { FirstLine = "420" }),
            new Run(
                new RunProperties(
                    new RunFonts
                    {
                        Ascii = "Calibri",
                        HighAnsi = "Calibri",
                        EastAsia = "宋体"
                    },
                    new Color { Val = "000000" },
                    new FontSize { Val = "21" }),
                new Text(text)));
    }

    private static Paragraph CreateCaption(string text)
    {
        return new Paragraph(
            new ParagraphProperties(
                new ParagraphStyleId { Val = "Caption" },
                new KeepLines()),
            new Run(new Text(text)));
    }

    private static Table CreateTableWithWidths(int[] widths)
    {
        var table = new Table();
        table.Append(new TableProperties(
            new TableWidth
            {
                Width = "5000",
                Type = TableWidthUnitValues.Pct
            },
            new TableBorders(
                new TopBorder
                {
                    Val = BorderValues.Single,
                    Size = 8,
                    Color = C.Primary
                },
                new LeftBorder { Val = BorderValues.Nil },
                new BottomBorder
                {
                    Val = BorderValues.Single,
                    Size = 8,
                    Color = C.Primary
                },
                new RightBorder { Val = BorderValues.Nil },
                new InsideHorizontalBorder
                {
                    Val = BorderValues.Single,
                    Size = 3,
                    Color = C.Border
                },
                new InsideVerticalBorder { Val = BorderValues.Nil }),
            new TableLayout { Type = TableLayoutValues.Fixed },
            new TableCellMarginDefault(
                new TopMargin { Width = "70", Type = TableWidthUnitValues.Dxa },
                new LeftMargin { Width = "100", Type = TableWidthUnitValues.Dxa },
                new BottomMargin { Width = "70", Type = TableWidthUnitValues.Dxa },
                new RightMargin { Width = "100", Type = TableWidthUnitValues.Dxa })));
        table.Append(new TableGrid(
            widths.Select(width => new GridColumn
            {
                Width = width.ToString()
            })));
        return table;
    }

    private static Table CreateBorderedTable(
        string[] headers,
        int[] widths,
        List<string[]> rows)
    {
        var table = new Table();
        table.Append(new TableProperties(
            new TableWidth
            {
                Width = "5000",
                Type = TableWidthUnitValues.Pct
            },
            new TableBorders(
                new TopBorder
                {
                    Val = BorderValues.Single,
                    Size = 4,
                    Color = "000000"
                },
                new LeftBorder
                {
                    Val = BorderValues.Single,
                    Size = 4,
                    Color = "000000"
                },
                new BottomBorder
                {
                    Val = BorderValues.Single,
                    Size = 4,
                    Color = "000000"
                },
                new RightBorder
                {
                    Val = BorderValues.Single,
                    Size = 4,
                    Color = "000000"
                },
                new InsideHorizontalBorder
                {
                    Val = BorderValues.Single,
                    Size = 4,
                    Color = "000000"
                },
                new InsideVerticalBorder
                {
                    Val = BorderValues.Single,
                    Size = 4,
                    Color = "000000"
                }),
            new TableLayout { Type = TableLayoutValues.Fixed },
            new TableCellMarginDefault(
                new TopMargin { Width = "50", Type = TableWidthUnitValues.Dxa },
                new LeftMargin { Width = "80", Type = TableWidthUnitValues.Dxa },
                new BottomMargin { Width = "50", Type = TableWidthUnitValues.Dxa },
                new RightMargin { Width = "80", Type = TableWidthUnitValues.Dxa })));
        table.Append(new TableGrid(
            widths.Select(width => new GridColumn { Width = width.ToString() })));

        var headerRow = new TableRow(new TableRowProperties(new TableHeader()));
        for (int i = 0; i < headers.Length; i++)
            headerRow.Append(CreatePlainCell(headers[i], widths[i], true));
        table.Append(headerRow);

        foreach (string[] row in rows)
        {
            var tableRow = new TableRow();
            for (int i = 0; i < row.Length; i++)
                tableRow.Append(CreatePlainCell(row[i], widths[i], false));
            table.Append(tableRow);
        }
        return table;
    }

    private static TableCell CreatePlainCell(
        string text,
        int width,
        bool bold)
    {
        var runProperties = new RunProperties(new RunFonts
        {
            Ascii = "Times New Roman",
            HighAnsi = "Times New Roman",
            EastAsia = "宋体"
        });
        if (bold) runProperties.Append(new Bold());
        runProperties.Append(new Color { Val = "000000" });
        runProperties.Append(new FontSize { Val = "20" });
        return new TableCell(
            new TableCellProperties(
                new TableCellWidth
                {
                    Width = width.ToString(),
                    Type = TableWidthUnitValues.Dxa
                },
                new TableCellVerticalAlignment
                {
                    Val = TableVerticalAlignmentValues.Center
                }),
            new Paragraph(
                new ParagraphProperties(
                    new SpacingBetweenLines
                    {
                        Before = "0",
                        After = "0",
                        Line = "260"
                    },
                    new Indentation { FirstLine = "0" }),
                new Run(runProperties, new Text(text))));
    }

    private static void AddInfoRow(Table table, string key, string value)
    {
        int[] widths = { 1900, 6400 };
        var row = new TableRow();
        row.Append(CreateCell(key, widths[0], true, C.Panel, "20"));
        row.Append(CreateCell(value, widths[1], false, C.White, "20"));
        table.Append(row);
    }

    private static TableCell CreateCell(
        string text,
        int width,
        bool bold,
        string fill,
        string fontSize)
    {
        var runProperties = new RunProperties(new RunFonts
        {
            Ascii = "Calibri",
            HighAnsi = "Calibri",
            EastAsia = "宋体"
        });
        if (bold) runProperties.Append(new Bold());
        runProperties.Append(new Color
        {
            Val = bold ? C.Primary : C.Dark
        });
        runProperties.Append(new FontSize { Val = fontSize });

        return new TableCell(
            new TableCellProperties(
                new TableCellWidth
                {
                    Width = width.ToString(),
                    Type = TableWidthUnitValues.Dxa
                },
                new Shading
                {
                    Val = ShadingPatternValues.Clear,
                    Fill = fill
                },
                new TableCellVerticalAlignment
                {
                    Val = TableVerticalAlignmentValues.Center
                }),
            new Paragraph(
                new ParagraphProperties(
                    new SpacingBetweenLines
                    {
                        Before = "0",
                        After = "0",
                        Line = "260"
                    },
                    new Indentation { FirstLine = "0" },
                    new Justification { Val = JustificationValues.Left }),
                new Run(runProperties, new Text(text))));
    }

    private static void AddInlineImage(
        Body body,
        MainDocumentPart main,
        string imagePath,
        string altText)
    {
        if (!File.Exists(imagePath))
            throw new FileNotFoundException($"Image not found: {imagePath}", imagePath);

        byte[] bytes = File.ReadAllBytes(imagePath);
        var imagePart = main.AddImagePart(ImagePartType.Png);
        using (var stream = new MemoryStream(bytes))
            imagePart.FeedData(stream);
        string imageId = main.GetIdOfPart(imagePart);

        if (bytes.Length < 24)
            throw new InvalidDataException($"Invalid PNG image: {imagePath}");
        int width = ReadBigEndianInt32(bytes, 16);
        int height = ReadBigEndianInt32(bytes, 20);
        if (width <= 0 || height <= 0)
            throw new InvalidDataException($"Invalid PNG dimensions: {imagePath}");

        const long maxWidthEmu = 145 * 360000L;
        long cx = maxWidthEmu;
        long cy = (long)(cx * ((double)height / width));
        uint id = _docPrId++;

        var drawing = new Drawing(new DW.Inline(
            new DW.Extent { Cx = cx, Cy = cy },
            new DW.EffectExtent
            {
                LeftEdge = 0,
                TopEdge = 0,
                RightEdge = 0,
                BottomEdge = 0
            },
            new DW.DocProperties
            {
                Id = id,
                Name = altText
            },
            new DW.NonVisualGraphicFrameDrawingProperties(
                new A.GraphicFrameLocks { NoChangeAspect = true }),
            new A.Graphic(new A.GraphicData(
                new PIC.Picture(
                    new PIC.NonVisualPictureProperties(
                        new PIC.NonVisualDrawingProperties
                        {
                            Id = id,
                            Name = $"{altText}.png"
                        },
                        new PIC.NonVisualPictureDrawingProperties()),
                    new PIC.BlipFill(
                        new A.Blip { Embed = imageId },
                        new A.Stretch(new A.FillRectangle())),
                    new PIC.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = 0, Y = 0 },
                            new A.Extents { Cx = cx, Cy = cy }),
                        new A.PresetGeometry
                        {
                            Preset = A.ShapeTypeValues.Rectangle
                        })))
            {
                Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
            }))
        {
            DistanceFromTop = 0,
            DistanceFromBottom = 0,
            DistanceFromLeft = 0,
            DistanceFromRight = 0
        });

        body.Append(new Paragraph(
            new ParagraphProperties(
                new KeepNext(),
                new SpacingBetweenLines { Before = "150", After = "60" },
                new Indentation { FirstLine = "0" },
                new Justification { Val = JustificationValues.Center }),
            new Run(drawing)));
    }

    private static int ReadBigEndianInt32(byte[] bytes, int offset)
    {
        return (bytes[offset] << 24)
            | (bytes[offset + 1] << 16)
            | (bytes[offset + 2] << 8)
            | bytes[offset + 3];
    }

    private static void AddFinalSection(
        Body body,
        string headerId,
        string footerId)
    {
        body.Append(new SectionProperties(
            new HeaderReference
            {
                Type = HeaderFooterValues.Default,
                Id = headerId
            },
            new FooterReference
            {
                Type = HeaderFooterValues.Default,
                Id = footerId
            },
            new PageSize
            {
                Width = (UInt32Value)(uint)A4W,
                Height = (UInt32Value)(uint)A4H,
                Orient = PageOrientationValues.Portrait
            },
            new PageMargin
            {
                Top = 1440,
                Right = 1800,
                Bottom = 1440,
                Left = 1800,
                Header = 720,
                Footer = 720,
                Gutter = 0
            }));
    }

    private static void SetUpdateFieldsOnOpen(MainDocumentPart main)
    {
        var settingsPart = main.DocumentSettingsPart
            ?? main.AddNewPart<DocumentSettingsPart>();
        settingsPart.Settings = new Settings(
            new UpdateFieldsOnOpen { Val = true });
    }
}
