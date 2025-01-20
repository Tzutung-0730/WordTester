using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordTester.Classes;

public class WordRender
{
    public static byte[] GenVocQuizPaper(string tmplPath, string[] vocList)
    {
        var tmplContent = File.ReadAllBytes(tmplPath);
        using (var ms = new MemoryStream())
        {
            ms.Write(tmplContent, 0, (int)tmplContent.Length);
            using (var doc = WordprocessingDocument.Open(ms, true))
            {
                var table = doc.MainDocumentPart.Document.Body.Elements<Table>().First();

                var firstRow = table.Elements<TableRow>().ElementAt(3);
                var firstCell = firstRow.Elements<TableCell>().ElementAt(1);



                var tbl = new Table();
                //設定框線
                var tp = new TableProperties(
                    //指定田字形六條線的樣式及線寬
                    new TableBorders(
                        //Size 單位為 1/8 點 [註]
                        new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                        new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                        new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                        new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                        new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 },
                        new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 8 }
                   )
                );
                tbl.AppendChild<TableProperties>(tp);
                var que = new Queue<string>(vocList);
                //A4 直式，兩欄
                while (que.Any())
                {
                    var tr = new TableRow();
                    //每一列放四欄
                    for (var i = 0; i < 2; i++)
                    {
                        var tc = new TableCell();
                        //第一欄為單字
                        tc.Append(new TableCellProperties(new TableCellWidth()
                        {
                            //寬度取15%
                            Type = TableWidthUnitValues.Pct,
                            Width = "15"
                        }));
                        var text = que.Any() ? que.Dequeue() : string.Empty;
                        tc.Append(new Paragraph(new Run(new Text(text))));
                        tr.Append(tc);
                        //第二欄為填空
                        tc = new TableCell();
                        tc.Append(new TableCellProperties(new TableCellWidth()
                        {
                            Type = TableWidthUnitValues.Pct,
                            Width = "35"
                        }));
                        tc.Append(new Paragraph(new Run(new Text(string.Empty))));
                        tr.Append(tc);
                    }
                    tbl.Append(tr);
                }
                doc.MainDocumentPart.Document.Body.Append(tbl);
            }
            return ms.ToArray();
        }
    }
}
