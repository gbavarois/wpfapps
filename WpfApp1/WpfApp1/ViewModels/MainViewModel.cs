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
        [NotifyCanExecuteChangedFor(nameof(PlaceRamCommand))]
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

        // 開いているタブのリスト
        [ObservableProperty]
        private ObservableCollection<DisplayEditorViewModel> _editorTabs = new();

        // 現在アクティブなタブ
        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(PlaceRamCommand))]
        [NotifyCanExecuteChangedFor(nameof(RemoveRamCommand))]
        private DisplayEditorViewModel? _activeTab;

        // 画面上に配置されているRAMアイテムのリスト（全タブ共通で管理する場合）
        [ObservableProperty]
        private ObservableCollection<RamItemViewModel> _ramItemVMList = new();

        // 現在開いているファイルのフルパスを保持する変数
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(WindowTitle))]
        public string _currentFilePath;

        // --- ダイアログ依頼イベント ---
        public event Func<Task<bool>>? RequestSaveBeforeContinue;
        public event Func<Task<string?>>? RequestOpenFilePath;
        public event Func<Task<string?>>? RequestSaveFilePath;
        public event Action<string>? SaveRequested;// --- Viewに保存実行させるイベント ---

        // タイトルバーに表示する文字列を合成
        public string WindowTitle => $"{(IsDirty ? "* " : "")}{(Path.GetFileName(_currentFilePath) ?? "無題")} - 画面ファイルエディタ";

        [ObservableProperty]
        private bool _isDraggingCatalog; // ドラッグ中かどうか

        public MainViewModel()
        {
            AddTab();
        }

        // [ファイル(F)] - [新規作成(N)]
        [RelayCommand]
        private async Task NewProject()
        {
            if (RequestSaveBeforeContinue != null)
            {
                var ok = await RequestSaveBeforeContinue.Invoke();
                if (!ok) return;
            }

            EditorTabs.Clear();
            CurrentFilePath = null;

            AddTabCommand.Execute(null);
            IsDirty = false;
        }

        // [ファイル(F)] - [開く(O)]
        [RelayCommand]
        private async Task OpenFile()
        {
            if (RequestSaveBeforeContinue != null)
            {
                var ok = await RequestSaveBeforeContinue.Invoke();
                if (!ok) return;
            }

            if (RequestOpenFilePath == null) return;

            var path = await RequestOpenFilePath.Invoke();
            if (string.IsNullOrEmpty(path)) return;

            LoadFromFile(path);
        }

        // [ファイル(F)] - [上書き保存(S)]
        [RelayCommand]
        private async Task SaveFile()
        {
            if (string.IsNullOrEmpty(CurrentFilePath))
            {
                if (RequestSaveFilePath == null) return;

                var path = await RequestSaveFilePath.Invoke();
                if (string.IsNullOrEmpty(path)) return;

                CurrentFilePath = path;
            }

            SaveRequested?.Invoke(CurrentFilePath);
            IsDirty = false;
        }

        // [ファイル(F)] - [名前を付けて保存(A)]
        [RelayCommand]
        private async Task SaveAsFile()
        {
            if (RequestSaveFilePath == null) return;

            var path = await RequestSaveFilePath.Invoke();
            if (string.IsNullOrEmpty(path)) return;

            CurrentFilePath = path;

            SaveRequested?.Invoke(path);
            IsDirty = false;
        }

        // [ファイル(F)] - [終了(X)]
        [RelayCommand]
        private void Exit()
        {
            Application.Current.Shutdown();
        }

        // [編集(E)] - [新しいタブ]
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

        // [編集(E)] - [RAMデータを配置]
        [RelayCommand(CanExecute = nameof(CanPlaceRam))]
        private void PlaceRam()
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
        private bool CanPlaceRam() => SelectedRamCatalog != null && ActiveTab != null;

        // [編集(E)] - [配置RAMを削除]
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

        // [編集(E)] - [RAMデータ読込]
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

            foreach (var tab in EditorTabs)
            {
                foreach (var ram in tab.PlacedRams)
                {
                    ram.RefreshCatalogInfo();
                }
            }
        }



        // JSONファイルからのデータ読み込みロジック
        private void LoadFromFile(string path)
        {
            var service = new JsonEditorService();
            var projectData = service.LoadFromJson(path);

            CurrentFilePath = path;
            EditorTabs.Clear();

            foreach (var tabData in projectData.Tabs)
            {
                var tabVM = new DisplayEditorViewModel(this)
                {
                    DisplayName = tabData.Title
                };

                foreach (var ram in tabData.Rams)
                {
                    tabVM.PlacedRams.Add(new RamItemViewModel(ram, this));
                }
                tabVM.RestoreData = tabData;
                EditorTabs.Add(tabVM);
            }

            ActiveTab = EditorTabs.FirstOrDefault();
            IsDirty = false;
        }

        public RamCatalog? GetCatalogFromAllSheets(string symbol)
        {
            return _ramTableMaster.FindCatalogBySymbol(symbol);
        }
    }
}
