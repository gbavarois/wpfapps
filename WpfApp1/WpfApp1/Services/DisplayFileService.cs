using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Documents;
using WpfApp1.Models;

namespace WpfApp1.Services
{
    public class DisplayFileService
    {
        IEnumerable<FormatData> _formatList;
        public DisplayFileService(IEnumerable<FormatData> formatList)
        {
            _formatList = formatList;
        }

        public void SaveToDisplayFile(ProjectSaveData data, string path)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var sjis = Encoding.GetEncoding("shift_jis");
            var jsonService = new JsonEditorService(); // 解析ロジック利用用

            using (var sw = new StreamWriter(path, false, sjis))
            {
                sw.WriteLine("X_OFF 1000");

                foreach (var tab in data.Tabs)
                {
                    sw.WriteLine(tab.Title.ToUpper());

                    // 1. XamlContent から一時的に FlowDocument を生成して解析
                    var doc = new FlowDocument();
                    var range = new TextRange(doc.ContentStart, doc.ContentEnd);
                    using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(tab.XamlContent)))
                    {
                        range.Load(ms, DataFormats.Xaml);
                    }

                    // 2. 解析して Lines と Colors をその場で取得
                    var lines = doc.Blocks.OfType<Paragraph>()
                                   .Select(p => new TextRange(p.ContentStart, p.ContentEnd).Text.Replace("\r", "").Replace("\n", ""))
                                   .ToList();
                    var colors = jsonService.ExtractColorInfo(doc);

                    // --- ブロック① & ②：ベーステキスト ---
                    for (int r = 0; r < lines.Count; r++)
                    {
                        string lineText = lines[r];
                        string lineBase = $"  0, {r,3}, &H04, \"{lineText}\"";

                        var ramsInRow = tab.Rams.Where(ram => ram.Row == r).ToList();
                        if (ramsInRow.Any())
                        {
                            StringBuilder sb = new StringBuilder(lineBase);
                            foreach (var ram in ramsInRow)
                            {
                                int offset = 1000 + ram.Column;
                                var fmt = _formatList.FirstOrDefault(f => f.Id == ram.FormatId)?.Code ?? "00";
                                sb.Append($" ,  {offset}, {fmt,3}, {ram.Address}");
                            }
                            sw.WriteLine(sb.ToString());
                        }
                        else { sw.WriteLine(lineBase); }
                    }

                    // --- ブロック③：色変え情報（ColumnBを使用） ---
                    foreach (var color in colors)
                    {
                        if (color.ColorIndex == "4") continue;

                        string fullLine = lines[color.Row];
                        // Substringもバイト数（半角単位）で切り出す処理が必要な場合はここを調整
                        // 現状のロジックを維持：
                        string subText = fullLine.Substring(color.Column, color.Length);

                        string colorHex = $"&H0{color.ColorIndex}";
                        // 外部アプリ用なので ColumnB を出力
                        sw.WriteLine($"{color.ColumnB,3}, {color.Row,3}, {colorHex}, \"{subText}\"");
                    }
                    sw.WriteLine($"  0,   0,   -1, \"\"");
                }
            }
        }

    }
}
