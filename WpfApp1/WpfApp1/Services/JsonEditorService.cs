using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
        public EditorSaveData CreateSaveData(RichTextBox rtb, IEnumerable<RamLayout> ramList)
        {
            var data = new EditorSaveData();

            // --- テキスト保存（行単位） ---
            foreach (var block in rtb.Document.Blocks)
            {
                if (block is Paragraph para)
                {
                    var text = new TextRange(para.ContentStart, para.ContentEnd).Text;

                    // 改行コード除去
                    text = text.Replace("\r", "").Replace("\n", "");

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
                    Symbol = ram.Catalog?.Symbol,
                    FormatId = ram.FormatId,
                    Offset = 0 // ← まだ未実装なら固定
                });
            }

            return data;
        }

        // JSON書き込み
        public void SaveToJson(EditorSaveData data, string path)
        {
            var json = System.Text.Json.JsonSerializer.Serialize(
                data,
                new System.Text.Json.JsonSerializerOptions
                {
                    WriteIndented = true
                });

            File.WriteAllText(path, json);
        }

        // JSON読み込み
        public EditorSaveData LoadFromJson(string path)
        {
            var json = File.ReadAllText(path);

            return System.Text.Json.JsonSerializer.Deserialize<EditorSaveData>(json);
        }

        // テキスト復元
        public void RestoreText(RichTextBox rtb, List<string> lines)
        {
            var doc = new FlowDocument();

            foreach (var line in lines)
            {
                var para = new Paragraph();

                // 1行は1RunでOK（後で色分割される）
                para.Inlines.Add(new Run(line));

                doc.Blocks.Add(para);
            }

            rtb.Document = doc;
        }

        // 色復元
        public void RestoreColors(RichTextBox rtb, List<TextColorInfo> colors)
        {
            ApplyColorInfo(rtb, colors);
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

                        // デフォルト（0）は保存しない
                        if (colorIndex != "0")
                        {
                            result.Add(new TextColorInfo
                            {
                                Row = row,
                                Column = col,
                                Length = text.Length,
                                ColorIndex = colorIndex
                            });
                        }

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

        public void ApplyColorInfo(RichTextBox rtb, List<TextColorInfo> list)
        {
            foreach (var info in list)
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
                    Catalog = catalog,              // nullの可能性あり
                    FormatSource = formatList,
                    FormatId = s.FormatId,
                    Row = s.Row,
                    Column = s.Column,
                    Symbol = s.Symbol
                };

                result.Add(ram);
            }

            return result;
        }
    }
}
