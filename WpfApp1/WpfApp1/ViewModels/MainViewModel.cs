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
        private Dictionary<string, List<RamCatalog>> _allSheets = new();

        [ObservableProperty] private ObservableCollection<string> _sheetNames = new();
        [ObservableProperty] private ObservableCollection<RamCatalog> _selectedSheetData = new();
        [ObservableProperty] private ObservableCollection<FormatData> _formatList = new();
        [ObservableProperty] private ObservableCollection<RamData> _ramdataList = new();

        [ObservableProperty] private List<RamCatalog> _allRamCatalogs = new();

        [ObservableProperty] private string? _selectedSheet;
        [ObservableProperty] private RamCatalog? _selectedRamCatalog;
        [ObservableProperty] private FormatData? _selectedFormat;
        [ObservableProperty] private RamData? _selectedRamdata;

        // シート選択が変わったら自動でリスト更新
        partial void OnSelectedSheetChanged(string? value)
        {
            SelectedSheetData.Clear();
            if (value != null && _allSheets.TryGetValue(value, out var data))
            {
                foreach (var d in data) SelectedSheetData.Add(d);
            }
        }

        // --- コマンドの実装 ---

        [RelayCommand]
        private void LoadExcel(string path)
        {
            _allSheets = ExcelLoader.Load(path);
            SheetNames.Clear();
            foreach (var name in _allSheets.Keys) SheetNames.Add(name);

            FormatList.Clear();
            var formats = ExcelLoader.LoadFormats(path);
            foreach (var f in formats) FormatList.Add(f);

            if (SheetNames.Count > 0) SelectedSheet = SheetNames[0];
        }

        [RelayCommand]
        private void AddRamFromCatalog()
        {
            if (SelectedRamCatalog == null) return;

            var newRam = new RamData
            {
                Catalog = SelectedRamCatalog,
                FormatId = SelectedRamCatalog.FormatId,
                // 初期座標などは必要に応じて
                Row = 0,
                Column = 0
            };
            newRam.ResolveReferences(FormatList, _allSheets.Values.SelectMany(x => x));
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
            var restoredRams = service.RestoreRamData(saveData.Rams, AllRamCatalogs, FormatList);

            RamdataList.Clear();
            foreach (var ram in restoredRams)
            {
                ram.ResolveReferences(FormatList, AllRamCatalogs);
                RamdataList.Add(ram);
            }

            return saveData; // View(RichTextBox)で使うためにデータを返す
        }
    }
}
