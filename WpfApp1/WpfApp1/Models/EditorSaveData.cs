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
        public string XamlContent { get; set; } = "";
        public List<RamLayout> Rams { get; set; } = new();
    }
}
