using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;

namespace DSPEditor.Helpers
{
    public static class RichTextHelper
    {
        // RichTextBox -> XAML文字列
        public static string GetXaml(RichTextBox rtb)
        {
            var range = new TextRange(rtb.Document.ContentStart, rtb.Document.ContentEnd);
            using (var ms = new MemoryStream())
            {
                range.Save(ms, DataFormats.Xaml);
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }

        // XAML文字列 -> RichTextBox
        public static void SetXaml(RichTextBox rtb, string xaml)
        {
            if (string.IsNullOrEmpty(xaml)) return;
            var range = new TextRange(rtb.Document.ContentStart, rtb.Document.ContentEnd);
            using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(xaml)))
            {
                range.Load(ms, DataFormats.Xaml);
            }
        }
    }
}
