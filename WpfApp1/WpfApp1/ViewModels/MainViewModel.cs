using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using WpfApp1.Models;
using WpfApp1.Services;

namespace WpfApp1.ViewModels
{
    public partial class MainViewModel : ObservableObject
    {
        private readonly RamTableMaster _model = new();
        // コンボボックス用のシート名リスト
        [ObservableProperty] private ObservableCollection<string> _sheetNames = new();
        // コンボボックスで選択されたページ名
        [ObservableProperty] private string? _selectedSheet;
        // DataGridに表示する現在のページのデータ
        [ObservableProperty] private ObservableCollection<RamCatalog> _currentCatalogs = new();
        //
        [ObservableProperty] private RamCatalog? _selectedRamCatalog;
        // FormatData用の一覧
        [ObservableProperty] private ObservableCollection<FormatData> _formatList = new();
        // 開いている全タブのリスト
        [ObservableProperty] private ObservableCollection<DisplayEditorViewModel> _editorTabs = new();
        // 現在アクティブなタブ
        [ObservableProperty] private DisplayEditorViewModel? _activeTab;

        [ObservableProperty] private ObservableCollection<RamItemViewModel> _ramItemVMList = new();

        [ObservableProperty] private FormatData? _selectedFormat;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(WindowTitle))] // タイトルも連動
        private bool _isDirty;

        // 現在開いているファイルのフルパスを保持する変数
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(WindowTitle))]
        public string _currentFilePath;

        // タイトルバーに表示する文字列を合成
        public string WindowTitle => $"{(IsDirty ? "* " : "")}{(Path.GetFileName(_currentFilePath) ?? "無題")} - 画面ファイルエディタ";

        // データを変更した時に呼ぶ
        public void MarkAsDirty() => IsDirty = true;

        public MainViewModel()
        {
            AddTab();
        }

        [RelayCommand]
        private void AddTab()
        {
            // 1. 候補となる文字のリストを定義（1～9, A～F）
            string candidates = "0123456789ABCDEF";

            // 2. 現在のタブ名から、使われている末尾の文字を抽出
            // 例: "Disp1" -> "1"
            var usedIds = EditorTabs
                .Select(t => t.DisplayName.Replace("Disp", ""))
                .ToList();

            // 3. 候補の中で、使われていない最初の文字を探す
            // FirstOrDefault で「条件に合う最初のもの」を取得
            char nextIdChar = candidates.FirstOrDefault(c => !usedIds.Contains(c.ToString()));

            // 4. 文字を決定（万が一15枚を超えた場合は、現在のカウントを振るなどの回避策）
            string nextId = nextIdChar != '\0' ? nextIdChar.ToString() : (EditorTabs.Count + 1).ToString();

            // 5. 新規タブ作成
            var tab = new DisplayEditorViewModel(this)
            {
                DisplayName = $"Disp{nextId}"
            };

            EditorTabs.Add(tab);
            ActiveTab = tab;
        }

        // SelectedSheetNameが変更されたときに自動で呼ばれるメソッド（Toolkitの機能）
        partial void OnSelectedSheetChanged(string value)
        {
            var data = _model.GetCatalogsBySheet(value);
            CurrentCatalogs.Clear();
            foreach (var item in data) CurrentCatalogs.Add(item);
            RefreshAllPlacedRams();
        }

        partial void OnSelectedRamCatalogChanged(RamCatalog? value)
        {
            AddRamFromCatalogCommand.NotifyCanExecuteChanged();
        }

        // アクティブなタブが変わった時に「配置・削除両方」を更新
        partial void OnActiveTabChanged(DisplayEditorViewModel? value)
        {
            AddRamFromCatalogCommand.NotifyCanExecuteChanged();
            RemoveRamCommand.NotifyCanExecuteChanged();
        }


        //private Dictionary<string, List<RamCatalog>> _allSheets = new();


        //[ObservableProperty] private ObservableCollection<RamCatalog> _selectedSheetData = new();


        //[ObservableProperty] private List<RamCatalog> _allRamCatalogs = new();



        //[ObservableProperty] private RamLayout? _selectedRamdata;               // これを消す
        //[ObservableProperty] private RamItemViewModel? _selectedRamItem;

        // シート選択が変わったら自動でリスト更新
        //partial void OnSelectedSheetChanged(string? value)
        //{
        //    SelectedSheetData.Clear();
        //    var count = _allSheets.Count;
        //    if (value != null && _allSheets.TryGetValue(value, out var data))
        //    {
        //        foreach (var d in data) SelectedSheetData.Add(d);
        //    }
        //}

        // --- コマンドの実装 ---
        //// 新規タブ追加コマンド
        //[RelayCommand]
        //private void AddNewTab()
        //{
        //    var newTab = new DisplayEditorViewModel(this) { DisplayNumber = $"Display {EditorTabs.Count + 1}" };
        //    EditorTabs.Add(newTab);
        //    ActiveTab = newTab; // 追加したタブを選択状態にする
        //}

        [RelayCommand]
        private void LoadExcel(string path)
        {
            _model.Load(path);

            SheetNames = new ObservableCollection<string>(_model.SheetNames);
            if (SheetNames.Count > 0) SelectedSheet = SheetNames[0];

            // ロードしたFormatDataを表示用リストにセット ---
            FormatList = new ObservableCollection<FormatData>(_model.Formats);
            RefreshAllPlacedRams();
        }

        // --- RAM配置コマンド ---
        [RelayCommand(CanExecute = nameof(CanAddRam))]
        private void AddRamFromCatalog()
        {
            if (SelectedRamCatalog == null || ActiveTab == null) return;

            var newRam = new RamLayout
            {
                Row = ActiveTab.CurrentRow,
                Column = ActiveTab.CurrentColumn,
                Offset = 0,
                Symbol = SelectedRamCatalog.Symbol,
                FormatId = SelectedRamCatalog.FormatId,
            };

            var vm = new RamItemViewModel(newRam, this);

            ActiveTab.PlacedRams.Add(vm);
            ActiveTab.SelectedRam = vm;
            IsDirty = true;
        }
        // カタログが選択されていて、かつタブが開いている時だけ有効
        private bool CanAddRam() => SelectedRamCatalog != null && ActiveTab != null;
        
        // --- RAM削除コマンド ---
        [RelayCommand(CanExecute = nameof(CanRemove))]
        private void RemoveRam()
        {
            if (ActiveTab?.SelectedRam != null)
            {
                ActiveTab.PlacedRams.Remove(ActiveTab.SelectedRam);
                ActiveTab.SelectedRam = null;
            }
        }
        private bool CanRemove() => ActiveTab?.SelectedRam != null;

        [RelayCommand]
        private void ClearSelection() => ActiveTab.SelectedRam = null;

        public void RefreshAllPlacedRams()
        {
            foreach (var tab in EditorTabs)
            {
                foreach (var ram in tab.PlacedRams)
                {
                    ram.RefreshCatalogInfo();
                }
            }
        }

        public RamCatalog? GetCatalogFromAllSheets(string symbol)
        {
            return _model.FindCatalogBySymbol(symbol);
        }
    }
}
