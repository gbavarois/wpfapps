using CommunityToolkit.Mvvm.ComponentModel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Collections.Generic;
using System.Linq;

namespace WpfApp1.Models
{
    public class RamLayout
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public int Offset { get; set; }
        public string Symbol { get; set; } = string.Empty;
        public string FormatId { get; set; } = string.Empty;
    }
}
