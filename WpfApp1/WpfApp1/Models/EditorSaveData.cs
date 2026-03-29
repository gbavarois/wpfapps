using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApp1.Models
{
    public class ProjectSaveData
    {
        // 各タブのデータのリスト
        public List<EditorData> Tabs { get; set; } = new();
    }

    public class EditorData
    {
        public string Title { get; set; } = "";
        public List<string> Lines { get; set; } = new();
        public List<TextColorInfo> Colors { get; set; } = new();
        public List<RamLayout> Rams { get; set; } = new();
    }
}
