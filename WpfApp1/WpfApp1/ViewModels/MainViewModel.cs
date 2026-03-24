using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
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



        [ObservableProperty] private ObservableCollection<RamLayout> _ramdataList = new();

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
        [ObservableProperty] private RamLayout? _selectedRamdata;

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
            if (SelectedRamCatalog == null) return;

            var newRam = new RamLayout
            {
                Catalog = SelectedRamCatalog,
                FormatId = SelectedRamCatalog.FormatId,
                // 初期座標などは必要に応じて
                Row = 0,
                Column = 0
            };
         //   newRam.ResolveReferences(FormatList, _allSheets.Values.SelectMany(x => x));
            RamdataList.Add(newRam);
            SelectedRamdata = newRam;
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

        public EditorSaveData? LoadFromJson(string path)
        {
            var service = new JsonEditorService();
            var saveData = service.LoadFromJson(path);

            // RAMデータの復元（ここはViewModelの仕事）
            var restoredRams = service.RestoreRamData(saveData.Rams, _currentCatalogs, FormatList);

            RamdataList.Clear();
            foreach (var ram in restoredRams)
            {
                ram.ResolveReferences(FormatList, _currentCatalogs);
                RamdataList.Add(ram);
            }

            return saveData; // View(RichTextBox)で使うためにデータを返す
        }
    }
}
