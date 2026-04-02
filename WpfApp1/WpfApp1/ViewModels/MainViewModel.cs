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
        // RAMテーブルのマスタークラス（Excelからのデータ読み込みと提供を担当）
        private readonly RamTableMaster _ramTableMaster = new();

        // ページ名とRAMカタログのリストを保持する辞書
        private Dictionary<string, ObservableCollection<RamCatalog>> _pages = new();
        
        // 現在選択されているページ名
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(CurrentCatalogs))]
        private string? _selectedPage;

        // ページ名のリスト（ComboBoxのItemsSourceにバインド）
        public IEnumerable<string> PageNames => _pages.Keys;

        // 現在選択されているページのRAMカタログアイテムのリスト（ListBoxのItemsSourceにバインド）
        public IEnumerable<RamCatalog> CurrentCatalogs =>
            SelectedPage != null && _pages.TryGetValue(SelectedPage, out var list)
                ? list
                : Enumerable.Empty<RamCatalog>();

        // DataGridで選択されたRAMカタログアイテム
        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(AddRamFromCatalogCommand))]
        private RamCatalog? _selectedRamCatalog;

        // フォーマットデータのリスト（これもExcelから読み込む）
        [ObservableProperty]
        private ObservableCollection<FormatData> _formatList = new();

        // 現在SelectedRamCatalogに基づいて、配置するRAMのフォーマットを特定するためのプロパティ
        [ObservableProperty]
        private FormatData? _selectedFormat;

        // データが変更されたかどうかを示すフラグ
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(WindowTitle))]
        private bool _isDirty;

        //// コンボボックスに表示するページ名のリスト
        //[ObservableProperty]
        //private ObservableCollection<string> _pageNames = new();

        //// 現在選択されているページ名
        //[ObservableProperty]
        //private string? _selectedPage;

        //// 現在表示しているページのRAMカタログアイテムのリスト
        //[ObservableProperty]
        //private ObservableCollection<RamCatalog> _currentCatalogs = new();

        //// フォーマットデータのリスト（これもExcelから読み込む）
        //[ObservableProperty]
        //private ObservableCollection<FormatData> _formatList = new();

        // 開いているタブのリスト
        [ObservableProperty]
        private ObservableCollection<DisplayEditorViewModel> _editorTabs = new();

        // 現在アクティブなタブ
        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(AddRamFromCatalogCommand))]
        [NotifyCanExecuteChangedFor(nameof(RemoveRamCommand))]
        private DisplayEditorViewModel? _activeTab;

        // 画面上に配置されているRAMアイテムのリスト（全タブ共通で管理する場合）
        [ObservableProperty]
        private ObservableCollection<RamItemViewModel> _ramItemVMList = new();

        

        // 現在開いているファイルのフルパスを保持する変数
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(WindowTitle))]
        public string _currentFilePath;

        // タイトルバーに表示する文字列を合成
        public string WindowTitle => $"{(IsDirty ? "* " : "")}{(Path.GetFileName(_currentFilePath) ?? "無題")} - 画面ファイルエディタ";

        [ObservableProperty]
        private bool _isDraggingCatalog; // ドラッグ中かどうか

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
            var usedIds = EditorTabs.Select(t => t.DisplayName.Replace("Disp", "")).ToList();

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

        // SelectedPageNameが変更されたときに自動で呼ばれるメソッド
        partial void OnSelectedPageChanged(string value)
        {
            //var data = _ramTableMaster.GetCatalogsByPage(value);
            //CurrentCatalogs.Clear();
            //foreach (var item in data) CurrentCatalogs.Add(item);
            RefreshAllPlacedRams();
        }

        //partial void OnSelectedRamCatalogChanged(RamCatalog? value)
        //{
        //    AddRamFromCatalogCommand.NotifyCanExecuteChanged();
        //}

        //// アクティブなタブが変わった時に「配置・削除両方」を更新
        //partial void OnActiveTabChanged(DisplayEditorViewModel? value)
        //{
        //    AddRamFromCatalogCommand.NotifyCanExecuteChanged();
        //    RemoveRamCommand.NotifyCanExecuteChanged();
        //}


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
            _ramTableMaster.Load(path);
            _pages = _ramTableMaster.Pages.ToDictionary(
                kvp => kvp.Key,
                kvp => new ObservableCollection<RamCatalog>(kvp.Value)
            );
            OnPropertyChanged(nameof(PageNames));   // ComboBox更新
            OnPropertyChanged(nameof(CurrentCatalogs)); // 念のため

            SelectedPage = PageNames.FirstOrDefault();

            FormatList = new ObservableCollection<FormatData>(_ramTableMaster.Formats);

            
            //PageNames = new ObservableCollection<string>(_ramTableMaster.PageNames);
            //if (PageNames.Count > 0) SelectedPage = PageNames[0];

            // ロードしたFormatDataを表示用リストにセット ---
            //FormatList = new ObservableCollection<FormatData>(_ramTableMaster.Formats);
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

        //[RelayCommand]
        //private void ClearSelection() => ActiveTab.SelectedRam = null;

		

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
            return _ramTableMaster.FindCatalogBySymbol(symbol);
        }
    }
}
