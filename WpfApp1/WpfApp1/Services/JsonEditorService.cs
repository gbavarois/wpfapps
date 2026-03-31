using DocumentFormat.OpenXml.Office2010.ExcelAc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Encodings.Web;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using WpfApp1.Helpers;
using WpfApp1.Models;

namespace WpfApp1.Services
{
    public class JsonEditorService
    {
        // 保存データ生成（RichText → DTO）
        public EditorData CreateSaveEditorData(string title, RichTextBox rtb, IEnumerable<RamLayout> ramList)
        {
            var data = new EditorData();

            // --- タイトル ---
            data.Title = title;

            // --- テキスト保存（行単位） ---
            foreach (var block in rtb.Document.Blocks)
            {
                if (block is Paragraph para)
                {
                    var text = new TextRange(para.ContentStart, para.ContentEnd).Text;
                    text = text.Replace("\r", "").Replace("\n", "");    // 改行コード除去
                    data.Lines.Add(text);
                }
            }

            // --- 色情報 ---
            data.Colors = ExtractColorInfo(rtb);

            // --- RAMデータ ---
            foreach (var ram in ramList)
            {
                data.Rams.Add(new RamLayout
                {
                    Row = ram.Row,
                    Column = ram.Column,
                    Offset = ram.Offset,
                    Symbol = ram.Symbol,
                    FormatId = ram.FormatId
                });
            }

            return data;
        }

        // JSON書き込み
        public void SaveToJson(ProjectSaveData data, string path)
        {
            var json = System.Text.Json.JsonSerializer.Serialize(
                data,
                new System.Text.Json.JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                });
            var sjis = System.Text.Encoding.GetEncoding("shift_jis");
            File.WriteAllText(path, json, sjis);
        }

        // JSON読み込み
        public ProjectSaveData LoadFromJson(string path)
        {
            var sjis = System.Text.Encoding.GetEncoding("shift_jis");
            var json = File.ReadAllText(path, sjis);

            return System.Text.Json.JsonSerializer.Deserialize<ProjectSaveData>(json);
        }

        // テキスト復元
        public void RestoreText(RichTextBox rtb, List<string> lines)
        {
            var doc = new FlowDocument();
            doc.PagePadding = rtb.Document.PagePadding;
            doc.PageWidth = rtb.Document.PageWidth;
            doc.ColumnWidth = rtb.Document.ColumnWidth;

            var paraStyle = rtb.FindResource(typeof(Paragraph)) as Style;
            foreach (var line in lines)
            {
                var para = new Paragraph();
                // スタイルを適用
                if (paraStyle != null)
                {
                    para.Style = paraStyle;
                }
                // 1行は1RunでOK（後で色分割される）
                para.Inlines.Add(new Run(line));

                doc.Blocks.Add(para);
            }

            rtb.Document = doc;
        }

        // 色復元
        public void RestoreColors(RichTextBox rtb, List<TextColorInfo> colors)
        {
            foreach (var info in colors)
            {
                var start = GetTextPointerAt(rtb, info.Row, info.Column);
                var end = GetTextPointerAt(rtb, info.Row, info.Column + info.Length);

                if (start != null && end != null)
                {
                    var range = new TextRange(start, end);
                    range.ApplyPropertyValue(
                        TextElement.ForegroundProperty,
                        new BrushConverter().ConvertFromString(ColorHelper.GetColorCode(info.ColorIndex))
                    );
                }
            }
        }

        // 内部：色情報抽出
        private List<TextColorInfo> ExtractColorInfo(RichTextBox rtb)
        {
            var result = new List<TextColorInfo>();

            int row = 0;
            foreach (var block in rtb.Document.Blocks)
            {
                if (block is not Paragraph para) continue;

                int col = 0;
                foreach (var inline in para.Inlines)
                {
                    if (inline is Run run)
                    {
                        string text = run.Text ?? "";
                        if (string.IsNullOrEmpty(text)) continue;

                        string colorIndex = ColorHelper.GetColorIndex(run.Foreground);
                        result.Add(new TextColorInfo
                        {
                            Row = row,
                            Column = col,
                            Length = text.Length,
                            ColorIndex = colorIndex
                        });
                        col += text.Length;
                    }
                }
                row++;
            }

            return result;
        }

        // 内部：TextPointer取得
        public TextPointer GetTextPointerAt(RichTextBox rtb, int row, int col)
        {
            int currentRow = 0;

            foreach (var block in rtb.Document.Blocks)
            {
                if (block is not Paragraph para) continue;

                if (currentRow == row)
                {
                    var pointer = para.ContentStart;
                    int currentCol = 0;

                    while (pointer != null)
                    {
                        if (pointer.GetPointerContext(LogicalDirection.Forward) == TextPointerContext.Text)
                        {
                            string text = pointer.GetTextInRun(LogicalDirection.Forward);

                            if (currentCol + text.Length >= col)
                            {
                                return pointer.GetPositionAtOffset(col - currentCol);
                            }

                            currentCol += text.Length;
                            pointer = pointer.GetPositionAtOffset(text.Length);
                        }
                        else
                        {
                            pointer = pointer.GetNextContextPosition(LogicalDirection.Forward);
                        }
                    }
                }

                currentRow++;
            }

            return null;
        }
        private bool BrushEquals(Brush a, Brush b)
        {
            if (a is SolidColorBrush sa && b is SolidColorBrush sb)
            {
                return sa.Color == sb.Color;
            }
            return false;
        }

        private string BrushToString(Brush brush)
        {
            if (brush is SolidColorBrush scb)
            {
                return scb.Color.ToString(); // "#FFRRGGBB"
            }
            return "#000000";
        }

        private Brush StringToBrush(string color)
        {
            return (Brush)new BrushConverter().ConvertFromString(color);
        }

        public List<RamLayout> RestoreRamData(List<RamLayout> saved,
                                            IEnumerable<RamCatalog> catalogList,
                                            IEnumerable<FormatData> formatList)
        {
            var result = new List<RamLayout>();

            foreach (var s in saved)
            {
                var catalog = catalogList.FirstOrDefault(c => c.Symbol == s.Symbol);

                var ram = new RamLayout
                {
                    //Catalog = catalog,              // nullの可能性あり
                    //FormatSource = formatList,
                    FormatId = s.FormatId,
                    Row = s.Row,
                    Column = s.Column,
                    Symbol = s.Symbol
                };

                result.Add(ram);
            }

            return result;
        }

        public void SaveProjectToJson(ProjectSaveData data, string path)
        {
            var json = System.Text.Json.JsonSerializer.Serialize(data, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(path, json);
        }

        public ProjectSaveData LoadProjectFromJson(string path)
        {
            var sjis = System.Text.Encoding.GetEncoding("shift_jis");
            var json = File.ReadAllText(path, sjis);
            return System.Text.Json.JsonSerializer.Deserialize<ProjectSaveData>(json);
        }
    }
}
