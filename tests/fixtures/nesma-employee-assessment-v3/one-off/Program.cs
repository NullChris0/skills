using System.Text.Json.Nodes;
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
    private const long A4WE = 7560000L;
    private const long A4HE = 10692000L;

    private static class C
    {
        public const string Primary = "344B46";
        public const string Secondary = "5E746D";
        public const string Accent = "B69563";
        public const string Dark = "24312E";
        public const string Mid = "58635F";
        public const string Light = "8A9792";
        public const string Border = "D7DDD9";
        public const string Panel = "F3F6F4";
        public const string Header = "DDE7E1";
        public const string White = "FFFFFF";
    }

    private static uint _docPrId = 1;

    public static void Main(string[] args)
    {
        string output = args.Length > 0
            ? args[0]
            : "/home/nc/projects/work/hx-employee-assessment-delivery-v3-new-v2/需求说明书.docx";
        string modelPath = "/tmp/hx-delivery-work/delivery-model.json";
        JsonNode root = JsonNode.Parse(File.ReadAllText(modelPath))!;
        JsonNode document = root["document"]!;

        using var package = WordprocessingDocument.Create(
            output,
            WordprocessingDocumentType.Document);
        var main = package.AddMainDocumentPart();
        main.Document = new Document(new Body());
        var body = main.Document.Body!;

        AddStyles(main);
        AddCoreProperties(package, document);
        AddNumbering(main);

        (string headerId, string footerId) = AddHeaderFooter(main, S(document["title"]));

        AddPlainCover(body, document);
        AddToc(body, document);
        AddBodyContent(body, main, document);
        AddFinalSection(body, footerId);
        SetUpdateFieldsOnOpen(main);

        main.Document.Save();
        package.Save();
        Console.WriteLine($"Generated: {output}");
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
                new FontSize { Val = "21" },
                new Color { Val = "000000" }))
        {
            Type = StyleValues.Paragraph,
            StyleId = "Normal",
            Default = true
        });

        styles.Append(CreateHeadingStyle(
            "Heading1", "heading 1", 0, "32", "000000", "0", "180", true));
        styles.Append(CreateHeadingStyle(
            "Heading2", "heading 2", 1, "28", "000000", "320", "120", false));
        styles.Append(CreateHeadingStyle(
            "Heading3", "heading 3", 2, "24", "000000", "280", "100", false));
        styles.Append(CreateHeadingStyle(
            "Heading4", "heading 4", 3, "22", "000000", "240", "80", false));
        styles.Append(CreateHeadingStyle(
            "Heading5", "heading 5", 4, "21", "000000", "200", "60", false));

        styles.Append(new Style(
            new StyleName { Val = "Caption" },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new Indentation { FirstLine = "0" },
                new SpacingBetweenLines { Before = "60", After = "300" }),
            new StyleRunProperties(
                new RunFonts
                {
                    Ascii = "Times New Roman",
                    HighAnsi = "Times New Roman",
                    EastAsia = "宋体"
                },
                new FontSize { Val = "20" },
                new Color { Val = "000000" }))
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
        string color,
        string before,
        string after,
        bool pageBreak)
    {
        var paragraphProperties = new StyleParagraphProperties(
            new KeepNext(),
            new KeepLines(),
            new SpacingBetweenLines { Before = before, After = after },
            new Indentation { FirstLine = "0" },
            new OutlineLevel { Val = level });
        if (pageBreak) paragraphProperties.Append(new PageBreakBefore());
        return new Style(
            new StyleName { Val = name },
            new BasedOn { Val = "Normal" },
            paragraphProperties,
            new StyleRunProperties(
                new Bold(),
                new FontSize { Val = size },
                new RunFonts
                {
                    Ascii = "Times New Roman",
                    HighAnsi = "Times New Roman",
                    EastAsia = "宋体"
                },
                new Color { Val = color }))
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
        var runProperties = new StyleRunProperties(
            new RunFonts
            {
                Ascii = "Times New Roman",
                HighAnsi = "Times New Roman",
                EastAsia = "宋体"
            },
            new Color { Val = "000000" },
            new FontSize { Val = "20" });
        if (bold) runProperties.Append(new Bold());
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

    private static void AddCoreProperties(WordprocessingDocument package, JsonNode document)
    {
        var props = package.PackageProperties;
        props.Title = S(document["title"]);
        props.Subject = "员工年度考核 v3-new-v2 需求说明书";
        props.Creator = "Codex";
        props.LastModifiedBy = "Codex";
        props.Created = new DateTime(2026, 10, 21);
        props.Modified = new DateTime(2026, 10, 21);
        props.Description = "基于 NESMA v3-new-v2 计数模型生成的正式需求交付件";
    }

    private static (string HeaderId, string FooterId) AddHeaderFooter(
        MainDocumentPart main,
        string title)
    {
        var headerPart = main.AddNewPart<HeaderPart>();
        string headerId = main.GetIdOfPart(headerPart);
        headerPart.Header = new Header(new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Right },
                new Indentation { FirstLine = "0" },
                new SpacingBetweenLines { After = "0" }),
            new Run(
                new RunProperties(
                    new RunFonts
                    {
                        Ascii = "Calibri",
                        HighAnsi = "Calibri",
                        EastAsia = "宋体"
                    },
                    new FontSize { Val = "16" },
                    new Color { Val = C.Light }),
                new Text(title))));

        var footerPart = main.AddNewPart<FooterPart>();
        string footerId = main.GetIdOfPart(footerPart);
        var paragraph = new Paragraph(new ParagraphProperties(
            new Justification { Val = JustificationValues.Center },
            new Indentation { FirstLine = "0" },
            new SpacingBetweenLines { Before = "0", After = "0" }));
        AppendFooterField(paragraph, " PAGE ", "1");
        footerPart.Footer = new Footer(paragraph);
        return (headerId, footerId);
    }

    private static void AppendFooterText(Paragraph paragraph, string text)
    {
        paragraph.Append(new Run(
            FooterRunProperties(),
            new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
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
            new FontSize { Val = "18" },
            new Color { Val = "000000" });
    }

    private static void AddCover(Body body, JsonNode document, string coverImageId)
    {
        body.Append(new Paragraph(
            new ParagraphProperties(
                new Indentation { FirstLine = "0" },
                new SpacingBetweenLines { Before = "0", After = "0" }),
            new Run(CreateFloatingBackground(coverImageId, _docPrId++, "CoverBackground"))));

        body.Append(new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "4050", After = "0" },
                new Indentation { Left = "1050", Right = "950", FirstLine = "0" }),
            CoverRun("员工年度考核", "96", true, C.White, 42)));

        body.Append(new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "0", After = "260" },
                new Indentation { Left = "1050", Right = "950", FirstLine = "0" }),
            CoverRun("需求说明书", "96", true, C.White, 42)));

        body.Append(new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "80", After = "180" },
                new Indentation { Left = "1050", Right = "950", FirstLine = "0" }),
            CoverRun(S(document["subtitle"]), "30", false, "DDE7E1", 18)));

        body.Append(new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "0", After = "1250" },
                new Indentation { Left = "1050", Right = "950", FirstLine = "0" }),
            CoverRun(S(document["version"]), "24", false, "F1D7A9", 12)));

        body.Append(new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "0", After = "0" },
                new Indentation { Left = "1050", Right = "950", FirstLine = "0" }),
            CoverRun(S(document["date"]) + " · 交付版", "22", false, "DDE7E1", 8)));

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
                },
                new TitlePage()))));
    }

    private static void AddPlainCover(Body body, JsonNode document)
    {
        body.Append(new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "6800", After = "120" },
                new Justification { Val = JustificationValues.Center },
                new Indentation { FirstLine = "0" }),
            new Run(
                new RunProperties(
                    new RunFonts
                    {
                        Ascii = "Times New Roman",
                        HighAnsi = "Times New Roman",
                        EastAsia = "宋体"
                    },
                    new Bold(),
                    new FontSize { Val = "32" },
                    new Color { Val = "000000" }),
                new Text(S(document["title"])))));

        body.Append(new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "0", After = "180" },
                new Justification { Val = JustificationValues.Center },
                new Indentation { FirstLine = "0" }),
            new Run(
                new RunProperties(
                    new RunFonts
                    {
                        Ascii = "Times New Roman",
                        HighAnsi = "Times New Roman",
                        EastAsia = "宋体"
                    },
                    new Bold(),
                    new FontSize { Val = "28" },
                    new Color { Val = "000000" }),
                new Text(S(document["subtitle"])))));

        body.Append(new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "180", After = "0" },
                new Justification { Val = JustificationValues.Center },
                new Indentation { FirstLine = "0" }),
            new Run(
                new RunProperties(
                    new RunFonts
                    {
                        Ascii = "Times New Roman",
                        HighAnsi = "Times New Roman",
                        EastAsia = "宋体"
                    },
                    new FontSize { Val = "22" },
                    new Color { Val = "000000" }),
                new Text(S(document["date"])))));

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

    private static Run CoverRun(
        string text,
        string textSize,
        bool bold,
        string color,
        int spacing)
    {
        return new Run(
            CoverRunProperties(textSize, bold, color, spacing),
            new Text(text));
    }

    private static RunProperties CoverRunProperties(
        string textSize,
        bool bold,
        string color,
        int spacing)
    {
        var properties = new RunProperties(
            new RunFonts
            {
                Ascii = "Calibri",
                HighAnsi = "Calibri",
                EastAsia = "微软雅黑"
            },
            new FontSize { Val = textSize },
            new Color { Val = color },
            new Spacing { Val = spacing });
        if (bold) properties.Append(new Bold());
        return properties;
    }

    private static void AddDocumentControl(Body body, JsonNode document)
    {
        body.Append(CreateHeading("文档控制", 1));
        body.Append(CreateParagraph(
            "本文是员工年度考核 v3-new-v2 的正式需求交付件。功能规模、数据功能、事务功能、类型和 FP 全部引用阶段一计数模型，不在本阶段重新计数。"));

        Table table = CreateTableWithWidths(new[] { 1800, 6500 });
        AddInfoRow(table, "文档名称", S(document["title"]));
        AddInfoRow(table, "交付版本", S(document["version"]));
        AddInfoRow(table, "计数口径", "NESMA 详细 FPA，新项目 ADD；13 个数据功能、37 个事务功能，267 UFP");
        AddInfoRow(table, "后端事实基线", "c761f4f529a80e2972436292f1aef35828fcf5a9");
        AddInfoRow(table, "前端事实基线", "bbdb9f02580432a4fd9225e92e5351cf3ada3610");
        AddInfoRow(table, "编制日期", S(document["date"]));
        body.Append(table);
        body.Append(CreateParagraph(
            "阅读约定：第七章描述数据模型，第八章描述页面业务功能，第九章逐项插入接口或主系统只读契约，附录提供 NESMA 覆盖矩阵。"));
    }

    private static void AddToc(Body body, JsonNode document)
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
                    new FontSize { Val = "34" },
                    new Color { Val = C.Primary }),
                new Text("目录"))));

        body.Append(new Paragraph(
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode(" TOC \\o \"1-4\" \\h \\z \\u ")
            {
                Space = SpaceProcessingModeValues.Preserve
            }),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate })));

        foreach (JsonNode section in document["sections"]!.AsArray())
        {
            int level = I(section["level"]);
            if (level >= 1 && level <= 4)
                body.Append(CreateTocEntry(
                    S(section["heading"]),
                    level,
                    S(section["page_hint"], "1")));
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
        JsonNode document)
    {
        foreach (JsonNode section in document["sections"]!.AsArray())
        {
            int level = I(section["level"]);
            body.Append(CreateHeading(S(section["heading"]), level));

            if (section["paras"] is JsonArray paragraphs)
            {
                foreach (JsonNode paragraph in paragraphs)
                    body.Append(CreateParagraph(S(paragraph)));
            }

            if (section["tables"] is JsonArray tables)
            {
                foreach (JsonNode tableSpec in tables)
                {
                    var headers = tableSpec["headers"]!.AsArray()
                        .Select(x => S(x))
                        .ToArray();
                    var widths = tableSpec["widths"]!.AsArray()
                        .Select(x => I(x))
                        .ToArray();
                    var rows = tableSpec["rows"]!.AsArray()
                        .Select(row => row!.AsArray().Select(x => S(x)).ToArray())
                        .ToList();
                    body.Append(CreateBorderedTable(headers, widths, rows));
                    body.Append(new Paragraph(new ParagraphProperties(
                        new SpacingBetweenLines { After = "100" })));
                }
            }

            if (section["meta"] is JsonArray metadata)
            {
                Table table = CreateTableWithWidths(new[] { 1900, 6400 });
                foreach (JsonNode item in metadata)
                {
                    if (item is JsonArray pair && pair.Count >= 2)
                        AddInfoRow(table, S(pair[0]), S(pair[1]));
                }
                body.Append(table);
                body.Append(new Paragraph(new ParagraphProperties(
                    new SpacingBetweenLines { After = "100" })));
            }

            if (section["images"] is JsonArray images)
            {
                foreach (JsonNode image in images)
                {
                    string path = S(image["path"]);
                    string caption = S(image["caption"]);
                    AddInlineImage(body, main, path, caption, 14.5);
                    if (!string.IsNullOrWhiteSpace(caption))
                        body.Append(CreateCaption(caption));
                }
            }
        }
    }

    private static void AddMappingAppendix(Body body, JsonArray mappings)
    {
        body.Append(CreateHeading("附录 A NESMA 覆盖矩阵", 1));
        body.Append(CreateParagraph(
            "下表逐项列出 13 个数据功能和 37 个事务功能的功能清单行、正文章节、接口或数据卡片以及覆盖状态。完整说明和筛选能力见同目录《映射表.xlsx》。"));

        int[] widths = { 900, 700, 2100, 480, 2500, 3600, 4100 };
        Table table = CreateTableWithWidths(widths);
        table.Append(CreateHeaderRow(
            widths,
            new[] { "NESMA", "类型", "名称", "FP", "功能清单行", "章节/页面/数据模型", "接口/数据卡片" }));

        foreach (JsonNode mapping in mappings)
        {
            table.Append(CreateBodyRow(
                widths,
                new[]
                {
                    S(mapping["id"]),
                    S(mapping["type"]),
                    S(mapping["name"]),
                    I(mapping["fp"]).ToString(),
                    S(mapping["feature_rows"]),
                    S(mapping["section_page_model"]),
                    S(mapping["interfaces"])
                },
                "15"));
        }
        body.Append(table);
    }

    private static Paragraph CreateHeading(string text, int level)
    {
        string style = level switch
        {
            <= 1 => "Heading1",
            2 => "Heading2",
            3 => "Heading3",
            4 => "Heading4",
            _ => "Heading5"
        };
        return new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = style }),
            new Run(new Text(text)));
    }

    private static Paragraph CreateParagraph(
        string text,
        string firstLine = "420",
        string color = "000000",
        string fontSize = "21")
    {
        return new Paragraph(
            new ParagraphProperties(
                new Indentation { FirstLine = firstLine },
                new SpacingBetweenLines
                {
                    After = "140",
                    Line = "360",
                    LineRule = LineSpacingRuleValues.Auto
                }),
            new Run(
                new RunProperties(
                    new RunFonts
                    {
                        Ascii = "Calibri",
                        HighAnsi = "Calibri",
                        EastAsia = "宋体"
                    },
                    new FontSize { Val = fontSize },
                    new Color { Val = color }),
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
            new TableLayout { Type = TableLayoutValues.Fixed },
            new TableBorders(
                new TopBorder
                {
                    Val = BorderValues.Single,
                    Size = 8,
                    Color = C.Primary
                },
                new BottomBorder
                {
                    Val = BorderValues.Single,
                    Size = 8,
                    Color = C.Primary
                },
                new LeftBorder { Val = BorderValues.Nil },
                new RightBorder { Val = BorderValues.Nil },
                new InsideHorizontalBorder
                {
                    Val = BorderValues.Single,
                    Size = 3,
                    Color = C.Border
                },
                new InsideVerticalBorder { Val = BorderValues.Nil }),
            new TableCellMarginDefault(
                new TopMargin { Width = "70", Type = TableWidthUnitValues.Dxa },
                new BottomMargin { Width = "70", Type = TableWidthUnitValues.Dxa },
                new LeftMargin { Width = "100", Type = TableWidthUnitValues.Dxa },
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
            new TableLayout { Type = TableLayoutValues.Fixed },
            new TableBorders(
                new TopBorder
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
                new LeftBorder
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
            new TableCellMarginDefault(
                new TopMargin { Width = "50", Type = TableWidthUnitValues.Dxa },
                new BottomMargin { Width = "50", Type = TableWidthUnitValues.Dxa },
                new LeftMargin { Width = "80", Type = TableWidthUnitValues.Dxa },
                new RightMargin { Width = "80", Type = TableWidthUnitValues.Dxa })));
        table.Append(new TableGrid(
            widths.Select(width => new GridColumn { Width = width.ToString() })));

        var headerRow = new TableRow(new TableRowProperties(new TableHeader()));
        for (int i = 0; i < headers.Length; i++)
            headerRow.Append(CreatePlainCell(headers[i], widths[i], true));
        table.Append(headerRow);

        foreach (var row in rows)
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
        var runProperties = new RunProperties(
            new RunFonts
            {
                Ascii = "Times New Roman",
                HighAnsi = "Times New Roman",
                EastAsia = "宋体"
            },
            new FontSize { Val = "20" },
            new Color { Val = "000000" });
        if (bold) runProperties.Append(new Bold());
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
                    new Indentation { FirstLine = "0" },
                    new SpacingBetweenLines
                    {
                        Before = "0",
                        After = "0",
                        Line = "260"
                    }),
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

    private static TableRow CreateHeaderRow(int[] widths, string[] cells)
    {
        var row = new TableRow(new TableRowProperties(new TableHeader()));
        for (int i = 0; i < cells.Length; i++)
            row.Append(CreateCell(cells[i], widths[i], true, C.Header, "16"));
        return row;
    }

    private static TableRow CreateBodyRow(
        int[] widths,
        string[] cells,
        string fontSize)
    {
        var row = new TableRow();
        for (int i = 0; i < cells.Length; i++)
            row.Append(CreateCell(cells[i], widths[i], false, C.White, fontSize));
        return row;
    }

    private static TableCell CreateCell(
        string text,
        int width,
        bool bold,
        string fill,
        string fontSize)
    {
        var runProperties = new RunProperties(
            new RunFonts
            {
                Ascii = "Calibri",
                HighAnsi = "Calibri",
                EastAsia = "宋体"
            },
            new FontSize { Val = fontSize },
            new Color { Val = bold ? C.Primary : C.Dark });
        if (bold) runProperties.Append(new Bold());

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
                    new Indentation { FirstLine = "0" },
                    new SpacingBetweenLines
                    {
                        Before = "0",
                        After = "0",
                        Line = "260"
                    },
                    new Justification { Val = JustificationValues.Left }),
                new Run(runProperties, new Text(text))));
    }

    private static void AddInlineImage(
        Body body,
        MainDocumentPart main,
        string imagePath,
        string altText,
        double maxWidthCm)
    {
        if (!File.Exists(imagePath))
        {
            Console.Error.WriteLine($"WARNING: image not found: {imagePath}");
            return;
        }

        var imagePart = main.AddImagePart(ImagePartType.Png);
        byte[] bytes = File.ReadAllBytes(imagePath);
        using (var stream = new MemoryStream(bytes))
            imagePart.FeedData(stream);
        string imageId = main.GetIdOfPart(imagePart);

        int width;
        int height;
        using (var stream = new MemoryStream(bytes))
        {
            stream.Seek(16, SeekOrigin.Begin);
            var widthBytes = new byte[4];
            var heightBytes = new byte[4];
            stream.Read(widthBytes, 0, 4);
            stream.Read(heightBytes, 0, 4);
            if (BitConverter.IsLittleEndian)
            {
                Array.Reverse(widthBytes);
                Array.Reverse(heightBytes);
            }
            width = BitConverter.ToInt32(widthBytes, 0);
            height = BitConverter.ToInt32(heightBytes, 0);
        }

        long cx = (long)(maxWidthCm * 360000L);
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
                new Justification { Val = JustificationValues.Center },
                new Indentation { FirstLine = "0" },
                new SpacingBetweenLines { Before = "150", After = "60" }),
            new Run(drawing)));
    }

    private static string AddImage(MainDocumentPart main, string path)
    {
        var imagePart = main.AddImagePart(ImagePartType.Png);
        using var stream = File.OpenRead(path);
        imagePart.FeedData(stream);
        return main.GetIdOfPart(imagePart);
    }

    private static Drawing CreateFloatingBackground(
        string imageId,
        uint docPrId,
        string name)
    {
        return new Drawing(new DW.Anchor(
            new DW.SimplePosition { X = 0, Y = 0 },
            new DW.HorizontalPosition(new DW.PositionOffset("0"))
            {
                RelativeFrom = DW.HorizontalRelativePositionValues.Page
            },
            new DW.VerticalPosition(new DW.PositionOffset("0"))
            {
                RelativeFrom = DW.VerticalRelativePositionValues.Page
            },
            new DW.Extent { Cx = A4WE, Cy = A4HE },
            new DW.EffectExtent
            {
                LeftEdge = 0,
                TopEdge = 0,
                RightEdge = 0,
                BottomEdge = 0
            },
            new DW.WrapNone(),
            new DW.DocProperties
            {
                Id = docPrId,
                Name = name
            },
            new DW.NonVisualGraphicFrameDrawingProperties(
                new A.GraphicFrameLocks { NoChangeAspect = true }),
            new A.Graphic(new A.GraphicData(
                new PIC.Picture(
                    new PIC.NonVisualPictureProperties(
                        new PIC.NonVisualDrawingProperties
                        {
                            Id = 0,
                            Name = $"{name}.png"
                        },
                        new PIC.NonVisualPictureDrawingProperties()),
                    new PIC.BlipFill(
                        new A.Blip { Embed = imageId },
                        new A.Stretch(new A.FillRectangle())),
                    new PIC.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = 0, Y = 0 },
                            new A.Extents { Cx = A4WE, Cy = A4HE }),
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
            DistanceFromRight = 0,
            SimplePos = false,
            RelativeHeight = 251658240,
            BehindDoc = true,
            Locked = false,
            LayoutInCell = true,
            AllowOverlap = true
        });
    }

    private static void AddSectionBreak(
        Body body,
        string headerId,
        string footerId,
        bool landscape)
    {
        uint width = landscape ? (uint)A4H : (uint)A4W;
        uint height = landscape ? (uint)A4W : (uint)A4H;
        body.Append(new Paragraph(new ParagraphProperties(
            new SectionProperties(
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
                    Width = (UInt32Value)width,
                    Height = (UInt32Value)height,
                    Orient = landscape
                        ? PageOrientationValues.Landscape
                        : PageOrientationValues.Portrait
                },
                new PageMargin
                {
                    Top = new Int32Value(landscape ? 1440 : 1800),
                    Right = new UInt32Value((uint)(landscape ? 1100 : 1440)),
                    Bottom = new Int32Value(1440),
                    Left = new UInt32Value((uint)(landscape ? 1100 : 1440)),
                    Header = 720,
                    Footer = 720,
                    Gutter = 0
                }))));
    }

    private static void AddFinalSection(Body body, string footerId)
    {
        body.Append(new SectionProperties(
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
                Top = new Int32Value(1440),
                Right = new UInt32Value((uint)1800),
                Bottom = new Int32Value(1440),
                Left = new UInt32Value((uint)1800),
                Header = 720,
                Footer = 720,
                Gutter = 0
            }));
    }

    private static void AddBackCover(
        Body body,
        JsonNode document,
        string backImageId)
    {
        body.Append(new Paragraph(
            new ParagraphProperties(
                new Indentation { FirstLine = "0" },
                new SpacingBetweenLines { Before = "0", After = "0" }),
            new Run(CreateFloatingBackground(backImageId, _docPrId++, "BackCoverBackground"))));

        body.Append(new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "5900", After = "300" },
                new Justification { Val = JustificationValues.Center },
                new Indentation { FirstLine = "0" }),
            CoverRun("员工年度考核", "54", true, C.Primary, 30)));

        body.Append(new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "0", After = "180" },
                new Justification { Val = JustificationValues.Center },
                new Indentation { FirstLine = "0" }),
            CoverRun(
                "需求说明书 · " + S(document["version"]),
                "24",
                false,
                C.Secondary,
                12)));

        body.Append(new Paragraph(
            new ParagraphProperties(
                new SpacingBetweenLines { Before = "0", After = "0" },
                new Justification { Val = JustificationValues.Center },
                new Indentation { FirstLine = "0" }),
            CoverRun("内部交付 · 2026年10月", "20", false, C.Light, 6)));

        body.Append(new SectionProperties(
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
                Footer = 0,
                Gutter = 0
            }));
    }

    private static void AddNumbering(MainDocumentPart main)
    {
        var numberingPart = main.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering = new Numbering(
            new AbstractNum(
                new Level(
                    new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new LevelText { Val = "%1." },
                    new LevelJustification { Val = LevelJustificationValues.Left },
                    new ParagraphProperties(new Indentation
                    {
                        Left = "720",
                        Hanging = "360",
                        FirstLine = "0"
                    }))
                {
                    LevelIndex = 0
                })
            {
                AbstractNumberId = 1
            },
            new NumberingInstance(
                new AbstractNumId { Val = 1 })
            {
                NumberID = 1
            });
    }

    private static void SetUpdateFieldsOnOpen(MainDocumentPart main)
    {
        var settingsPart = main.DocumentSettingsPart ??
            main.AddNewPart<DocumentSettingsPart>();
        settingsPart.Settings = new Settings(
            new UpdateFieldsOnOpen { Val = true },
            new DisplayBackgroundShape());
    }
}
