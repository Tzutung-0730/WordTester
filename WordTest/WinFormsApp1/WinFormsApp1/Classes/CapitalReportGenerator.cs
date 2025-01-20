using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Color = DocumentFormat.OpenXml.Wordprocessing.Color;

namespace WordTester.Classes;

public class CapitalWordRender
{
    // 懶人設定
    // 列舉 Word 的設定行為
    [Flags]
    private enum CellRender
    {
        None = 0,
        VerticalCenter = 1,
        Text_TopToBottomRightToLeftRotated = 2,
        MergeCell = 4,
        CheckTopBorder = 8,
    }

    private enum ParagraphRender
    { 
    }



    public static void TestGenCapitalPlanYear(string source, string target, string planYear)
    {
        GenerateActionPlanYearReport(source, target, planYear, StaticModel.ActionPlanReport);
    }

    public static void TestGenCapitalPlanQuarter(string source, string target, string planYear)
    {
        GenerateActionPlanQuarterReport(source, target, planYear, StaticModel.ActionPlanReport.Where(x=>x.SeasonSeqNo == 1));
    }

    /// <summary>
    /// 產生永續計畫年度報告
    /// </summary>
    /// <param name="templatePath"></param>
    /// <param name="outputFilePath"></param>
    /// <param name="title"></param>
    /// <param name="datas"></param>
    private static void GenerateActionPlanYearReport(string templatePath, string outputFilePath, string title, IEnumerable<ActionPlanReport> datas)
    {
        var tmpContent = File.ReadAllBytes(templatePath);

        using var ms = new MemoryStream();
        ms.Write(tmpContent, 0, tmpContent.Length);

        using (var doc = WordprocessingDocument.Open(ms, true))
        {
            var titlePara = doc.MainDocumentPart.Document.Body.Elements<Paragraph>();
            titlePara.First().Elements<Run>().First().Elements<Text>().First().Text = title;

            // 取得表格
            var table = doc.MainDocumentPart.Document.Body.Elements<Table>().First();

            // 對表格的第五欄至第八欄中的年度進行改變 (4..8) 代表從第五欄開始到第八欄
            var titleRow = table.Elements<TableRow>().First();
            titleRow.Elements<TableCell>().ToList()[4..8].ForEach(x =>
            {
                x.Elements<Paragraph>().First().Elements<Run>().First().Elements<Text>().First().Text = title;
            });

            // 群組化資料
            var groupByIssue = datas.GroupBy(x => new
            {
                x.IssueSeqNo,
                x.IssueName,
                x.StakeHolders,
                x.PlanSeqNo,
                x.PlanName
            })
            .Select(x => new
            {
                x.Key.IssueSeqNo,
                x.Key.IssueName,
                x.Key.StakeHolders,
                x.Key.PlanSeqNo,
                x.Key.PlanName,
                Seasons = x.GroupBy(y => new { y.SeasonSeqNo })
                            .Select(z => new
                            {
                                z.Key.SeasonSeqNo,
                                Issues = z.Select(zz => new
                                {
                                    zz.SectionNo,
                                    zz.ActionPlan
                                })
                            })
            });

            List<object> cellContextChange = [null, null, null, null];

            // OpenOffice 的 Twips（1 公分 = 567 Twips）
            // 依據 datas 的資料，產生表格內容
            foreach (var data in groupByIssue)
            {
                var tr = new TableRow();

                TableRowProperties tableProperties = new(new CantSplit()); // 表示允許跨頁分隔線
                tr.Append(tableProperties);

                var tc = CellVerticalCenter();

                tc.Append(new Paragraph(SetParagraphAlignCenter(), new Run(SetFont(), new Text(data.IssueSeqNo.ToString()))));
                tc.Append(CheckMergeMethod(cellContextChange, data.IssueSeqNo, 0));
                tc.Append(new TableCellWidth() { Width = "567", Type = TableWidthUnitValues.Dxa });
                tr.Append(tc);

                tc = CellVerticalCenter();
                tc.Append(new Paragraph(SetParagraphAlignCenter(), new Run(SetFont(), new Text(data.IssueName))));
                tc.Append(CheckMergeMethod(cellContextChange, data.IssueName, 1));
                tc.Append(new TableCellWidth() { Width = (2.4 * 567).ToString(), Type = TableWidthUnitValues.Dxa });
                tr.Append(tc);

                tc = CellVerticalCenter();
                foreach (var s in data.StakeHolders.Split("、"))
                    tc.Append(new Paragraph(SetParagraphAlignCenter(), new Run(SetFont(), new Text(s))));
                tc.Append(new TableCellWidth() { Width = (2.4 * 567).ToString(), Type = TableWidthUnitValues.Dxa });
                tr.Append(tc);

                tc = new TableCell();
                tc.Append(new Paragraph(SetParagraphFirstLine(data.PlanSeqNo.ToString().Length+1), new Run(SetFont(), new Text($"{data.PlanSeqNo}.{data.PlanName}"))));
                tc.Append(new TableCellWidth() { Width = (4.5 * 567).ToString(), Type = TableWidthUnitValues.Dxa });
                tr.Append(tc);

                // 放入四季的資料
                foreach (var season in data.Seasons)
                {
                    tc = new TableCell();

                    foreach (var issue in season.Issues)
                    {
                        var chapterNumber = "";
                        var chapterNumberLength = 0;

                        if (issue.SectionNo != null && issue.ActionPlan != null)
                        {
                            chapterNumber = $"{data.PlanSeqNo}.{issue.SectionNo}{issue.ActionPlan}";
                            chapterNumberLength = $"{data.PlanSeqNo}.{issue.SectionNo}".Length;
                        }

                        tc.Append(new Paragraph(SetParagraphFirstLine(3), new Run(SetFont(), new Text(chapterNumber))));
                        tc.Append(new TableCellWidth() { Width = (4.6 * 567).ToString(), Type = TableWidthUnitValues.Dxa });
                    }

                    tr.Append(tc);
                }

                table.Append(tr);
            }
        }
        File.WriteAllBytes(outputFilePath, ms.ToArray());
    }

    private static void GenerateActionPlanQuarterReport(string templatePath, string outputFilePath, string title, IEnumerable<ActionPlanReport> datas)
    {
        var tmpContent = File.ReadAllBytes(templatePath);
        using var ms = new MemoryStream();
        ms.Write(tmpContent, 0, tmpContent.Length);
        using (var doc = WordprocessingDocument.Open(ms, true))
        {
            // 取得表格
            var table = doc.MainDocumentPart.Document.Body.Elements<Table>().First();

            // 變更抬頭文字
            var titleRow = table.Elements<TableRow>().ToList();
            var titleCell = titleRow[0].Elements<TableCell>().First();
            titleCell.Elements<Paragraph>().First().Elements<Run>().First().Elements<Text>().First().Text = title;

            // 對表格的第五欄至第八欄中的年度進行改變 (4..6) 代表從第五欄開始到第六欄
            titleRow[1].Elements<TableCell>().ToList()[4..6].ForEach(x =>
            {
                x.Elements<Paragraph>().First().Elements<Run>().First().Elements<Text>().First().Text = title;
            });

            List<object> cellContextChange = [null, null, null, null];

            // OpenOffice 的 Twips（1 公分 = 567 Twips）
            // 依據 datas 的資料，產生表格內容
            foreach (var data in datas)
            {
                var tr = new TableRow();

                TableRowProperties tableProperties = new(new CantSplit()); // 表示允許跨頁分隔線
                tr.Append(tableProperties);

                var tc = new TableCell();

                // 項目
                var verticalMerge = CheckVerticalMergeMethod(cellContextChange, data.IssueSeqNo, 0);
                tr.Append(
                    new TableCell(
                        SetTableCellProperties(CellRender.VerticalCenter | CellRender.MergeCell, 0.6, verticalMerge),
                        new Paragraph(
                            SetParagraphAlignCenter(),
                            new Run(
                                SetFont(),
                                new Text(data.IssueSeqNo.ToString())
                            )
                        )
                    )
                );

                // 重大性議題
                verticalMerge = CheckVerticalMergeMethod(cellContextChange, data.IssueName, 1);
                tr.Append(
                    new TableCell(
                        SetTableCellProperties(CellRender.VerticalCenter | CellRender.Text_TopToBottomRightToLeftRotated | CellRender.MergeCell, 1.5, verticalMerge),
                        new Paragraph(
                            SetParagraphAlignCenter(),
                            new Run(
                                SetFont(),
                                new Text(data.IssueName)
                            )
                        )
                    )
                );

                // 利害關係人
                verticalMerge = CheckVerticalMergeMethod(cellContextChange, data.StakeHolders, 2);
                tc = new TableCell();
                tc.Append(
                    SetTableCellProperties(CellRender.VerticalCenter | CellRender.MergeCell, 2.1, verticalMerge)
                );
                foreach (var s in data.StakeHolders.Split("、"))
                    tc.Append(new Paragraph(SetParagraphAlignCenter(), new Run(SetFont(), new Text(s))));
                tc.Append();
                tr.Append(tc);

                // 預計推動計畫與方向
                verticalMerge = CheckVerticalMergeMethod(cellContextChange, $"{data.PlanSeqNo}.{data.PlanName}", 3);
                tc = new TableCell();
                tc.Append(
                    SetTableCellProperties(CellRender.MergeCell, 3.0, verticalMerge),
                    new Paragraph(
                        SetParagraphFirstLine(data.PlanSeqNo.ToString().Length + 1),
                        new Run(
                            SetFont(),
                            new Text($"{data.PlanSeqNo}.{data.PlanName}")
                        )
                    )
                );
                tr.Append(tc);

                var sectionId = $"{data.PlanSeqNo}.{data.SectionNo}";

                // 某季的推動計畫
                tr.Append(new TableCell(
                    SetTableCellProperties(CellRender.CheckTopBorder, 4.5, verticalMerge),
                    new Paragraph(
                        SetParagraphFirstLine(sectionId.Length),
                        new Run(SetFont(),
                            new Text($"{sectionId}{data.ActionPlan}")
                        ),
                        new Break(),
                        new Run(
                            SetFont(new Color() { Val = "FF0000" }),
                            new Text($"({data.EstimatedDelivery})")
                        )
                    )
                ));


                // 某季的進度符合預期
                tc = new TableCell(
                    SetTableCellProperties(CellRender.CheckTopBorder, 4.5, verticalMerge)
                );

                if (string.IsNullOrEmpty(data.OnTrackDesc) == false)
                {
                    tc.Append(
                        new Paragraph(
                            SetParagraphFirstLine(sectionId.Length),
                            new Run(SetFont(),
                                new Text($"{sectionId}{data.OnTrackDesc}")
                            ),
                            new Break(),
                            new Run(
                                SetFont(new Color() { Val = "FF0000" }),
                                new Text($"({data.ActualDelivery})")
                            )
                        )
                    );
                }
                else
                {
                    // 沒有 Data 要給空的段落
                    tc.Append(
                        new Paragraph(
                            new Run(
                                new Text("")
                            )
                        )
                    );
                }
                tr.Append(tc);

                // 某季的進度未符合預期
                tc = new TableCell(
                    SetTableCellProperties(CellRender.CheckTopBorder, 4.5, verticalMerge)
                );
                if (string.IsNullOrEmpty(data.OffTrackDesc) == false)
                {
                    tc.Append(
                        new Paragraph(
                            SetParagraphFirstLine(sectionId.Length),
                            new Run(SetFont(),
                                new Text($"{sectionId}{data.OffTrackDesc}")
                            ),
                            new Break(),
                            new Run(SetFont(new Color() { Val = "FF0000" }),
                                new Text($"({data.ActualDelivery})")
                            )
                        )
                    );
                }
                else
                {
                    // 沒有 Data 要給空的段落
                    tc.Append(
                        new Paragraph(
                            new Run(
                                new Text("")
                            )
                        )
                    );
                }
                tr.Append(tc);

                table.Append(tr);
            }

            foreach(var tableCell in table.Elements<TableRow>().Last().Elements<TableCell>())
            {
                tableCell.Elements<TableCellProperties>().ToList().ForEach(x => x.Append(new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 16 }));
            }
        }
        File.WriteAllBytes(outputFilePath, ms.ToArray());
    }

    private static TableCellProperties SetTableCellProperties(CellRender render, double width, MergedCellValues? mergedCellValues = null)
    {
        var cellProperties = new TableCellProperties();

        cellProperties.Append(new TableCellWidth() { Width = (width * 567).ToString(), Type = TableWidthUnitValues.Dxa });

        if (render.HasFlag(CellRender.VerticalCenter))
            cellProperties.Append(new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center });

        if (render.HasFlag(CellRender.Text_TopToBottomRightToLeftRotated))
            cellProperties.Append(new TextDirection() { Val = TextDirectionValues.TopToBottomRightToLeftRotated });

        if (render.HasFlag(CellRender.MergeCell) && mergedCellValues != null)
            cellProperties.Append(new VerticalMerge() { Val = mergedCellValues });

        if (render.HasFlag(CellRender.CheckTopBorder) &&
            mergedCellValues != null &&
            mergedCellValues == MergedCellValues.Continue)
        {
            cellProperties.Append(
                new TableCellBorders(
                    //Size 單位為 1/8 點 [註]
                    new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Nil) },
                    new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Nil) },
                    new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 16 },
                    new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 16 }
                )
            );
        }
        else
        {
            cellProperties.Append(
                new TableCellBorders(
                    //Size 單位為 1/8 點 [註]
                    new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 16 },
                    new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Nil) },
                    new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 16 },
                    new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 16 }
                )
            );
        }

        return cellProperties;
    }



    private static TableCellWidth SetTableCellWidth(double width)
    {
        return new TableCellWidth() { Width = (width * 567).ToString(), Type = TableWidthUnitValues.Dxa };
    }

    private static TableCellBorders SetTableCellBorder()
    {
        //指定田字形六條線的樣式及線寬
        return new TableCellBorders(
                    //Size 單位為 1/8 點 [註]
                    new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 16 },
                    new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 16 },
                    new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 16 },
                    new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 16 },
                    new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 16 },
                    new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 16 }
                );
    }

    private static MergedCellValues CheckVerticalMergeMethod(List<object> cellContextChange, object data, int index)
    {
        // 確認是否要合併儲存格
        if (cellContextChange[index] == null ||
            cellContextChange[index].ToString() != data.ToString())
        {
            cellContextChange[index] = data;
            return MergedCellValues.Restart;
        }
        else
        {
            return MergedCellValues.Continue;
        }
    }

    private static TableCellProperties CheckMergeMethod(List<object> cellContextChange, object data, int index)
    {
        TableCellProperties cellProp;

        // 確認是否要合併儲存格
        if (cellContextChange[index] == null ||
            cellContextChange[index].ToString() != data.ToString())
        {
            cellProp = new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart });
            cellContextChange[index] = data;
        }
        else
        {
            cellProp = new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Continue });
        }

        return cellProp;
    }

    /// <summary>
    /// 欄位垂直置中
    /// </summary>
    private static TableCell CellVerticalCenter()
    {
        var cellProperties = new TableCellProperties();
        cellProperties.Append(new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center });
        return new TableCell(cellProperties);
    }

    /// <summary>
    /// 段落置中
    /// </summary>
    private static Justification ParagraphCenter()
    {
        return new Justification() { Val = JustificationValues.Center };
    }

    /// <summary>
    /// 首字有幾個半型字元
    /// </summary>
    /// <param name="chars"></param>
    /// <returns></returns>
    private static ParagraphProperties SetParagraphFirstLine(int chars) 
    {
        var paragraphProperties = new ParagraphProperties();
        var ident = new Indentation() { HangingChars = chars*55 };
        var para = new Justification() { Val = JustificationValues.Both };
        paragraphProperties.Append(ident, para);

        return paragraphProperties;
    }

    private static ParagraphProperties SetParagraphAlignCenter()
    {
        var paragraphProperties = new ParagraphProperties();
        paragraphProperties.Append(ParagraphCenter());

        return paragraphProperties;
    }

    private static RunProperties SetFont(Color? color = null)
    {
        var runproperty = new RunProperties();
        runproperty.Append(new RunFonts() { Ascii = "標楷體", HighAnsi = "標楷體", EastAsia = "標楷體" });

        if (color != null)
        {
            runproperty.Append(color);
        }

        return runproperty;
    }

}
