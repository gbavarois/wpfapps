using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApp1.Models
{
    public class EditorSaveData
    {
        public List<string> Lines { get; set; } = new();
        public List<TextColorInfo> Colors { get; set; } = new();
        public List<RamSaveData> Rams { get; set; } = new();
    }
}
