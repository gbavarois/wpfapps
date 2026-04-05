using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WpfApp1.Models;
using WpfApp1.ViewModels;

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
            // SJISエンコーディングの登録（.NET Core/5+の場合に必要）
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var sjis = Encoding.GetEncoding("shift_jis");

            using (var sw = new StreamWriter(path, false, sjis))
            {
                // 0. 共通ヘッダー（相対桁1000対応）
                sw.WriteLine("X_OFF 1000");

                foreach (var tab in data.Tabs)
                {
                    // 1. 各タブの先頭に DISPLAY*
                    sw.WriteLine(tab.Title.ToUpper());

                    // --- ブロック① & ②：ベーステキストと配置RAM ---
                    for (int r = 0; r < tab.Lines.Count; r++)
                    {
                        string lineText = tab.Lines[r];
                        string lineBase = $"  0, {r,3}, &H04, \"{lineText}\"";

                        var ramsInRow = tab.Rams.Where(ram => ram.Row == r).ToList();

                        if (ramsInRow.Any())
                        {
                            StringBuilder sb = new StringBuilder(lineBase);
                            foreach (var ram in ramsInRow)
                            {
                                int offset = 1000 + ram.Column;
                                string addr = ram.Address;
                                var fmt = _formatList.FirstOrDefault(f => f.Id == ram.FormatId)?.Code ?? "00";

                                sb.Append($" ,  {offset}, {fmt,3}, {addr}");
                            }
                            sw.WriteLine(sb.ToString());
                        }
                        else
                        {
                            sw.WriteLine(lineBase);
                        }
                    }

                    // --- ブロック③：色変え情報 ---
                    foreach (var color in tab.Colors)
                    {
                        if (color.ColorIndex == "4") continue;

                        string fullLine = tab.Lines[color.Row];
                        string subText = "";

                        if (color.Column < fullLine.Length)
                        {
                            int len = Math.Min(color.Length, fullLine.Length - color.Column);
                            subText = fullLine.Substring(color.Column, len);
                        }

                        string colorHex = $"&H0{color.ColorIndex}";
                        sw.WriteLine($"{color.Column,3}, {color.Row,3}, {colorHex}, \"{subText}\"");
                    }
                }
            }
        }

    }
}
