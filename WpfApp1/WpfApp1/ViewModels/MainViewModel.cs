using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
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


        [ObservableProperty] private ObservableCollection<RamLayout> _ramdataList = new();          // これを消す

        [ObservableProperty] private ObservableCollection<RamItemViewModel> _ramItemVMList = new();

        public MainViewModel()
        {
            AddTab();
        }

        [RelayCommand]
        private void AddTab()
        {
            var tab = new DisplayEditorViewModel(this)
            {
                DisplayNumber = $"Disp{EditorTabs.Count}"
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
        }



        //private Dictionary<string, List<RamCatalog>> _allSheets = new();


        //[ObservableProperty] private ObservableCollection<RamCatalog> _selectedSheetData = new();
        

        //[ObservableProperty] private List<RamCatalog> _allRamCatalogs = new();



        [ObservableProperty] private FormatData? _selectedFormat;
        [ObservableProperty] private RamLayout? _selectedRamdata;               // これを消す
        [ObservableProperty] private RamItemViewModel? _selectedRamItem;

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
        // 新規タブ追加コマンド
        [RelayCommand]
        private void AddNewTab()
        {
            var newTab = new DisplayEditorViewModel(this) { DisplayNumber = $"Display {EditorTabs.Count + 1}" };
            EditorTabs.Add(newTab);
            ActiveTab = newTab; // 追加したタブを選択状態にする
        }

        [RelayCommand]
        private void LoadExcel(string path)
        {
            _model.Load(path);

            SheetNames = new ObservableCollection<string>(_model.SheetNames);
            if (SheetNames.Count > 0) SelectedSheet = SheetNames[0];

            // ロードしたFormatDataを表示用リストにセット ---
            FormatList = new ObservableCollection<FormatData>(_model.Formats);
        }

        [RelayCommand]
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
            //SelectedRamItem = vm; // これは一旦保留

        }

        [RelayCommand(CanExecute = nameof(CanRemove))]
        private void RemoveRam()
        {
            if (SelectedRamdata != null)
            {
                RamdataList.Remove(SelectedRamdata);
                SelectedRamdata = null;
            }
        }
        private bool CanRemove() => SelectedRamdata != null;

        [RelayCommand]
        private void ClearSelection() => SelectedRamdata = null;
    }
}
