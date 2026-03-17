using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using WpfApp1.Models;
using WpfApp1.Services;

namespace WpfApp1.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        Dictionary<string, List<RamCatalog>> allSheets = new();
        public ObservableCollection<string> SheetNames { get; } = new();
        public ObservableCollection<RamCatalog> SelectedSheetData { get; } = new();
        
        private string _selectedSheet;

        public string SelectedSheet
        {
            get => _selectedSheet;
            set
            {
                _selectedSheet = value;
                OnPropertyChanged();
                UpdateSheet();
            }
        }

        private RamCatalog _selectedRamCatalog; // 型を RamCatalog に修正
        public RamCatalog SelectedRamCatalog
        {
            get => _selectedRamCatalog;
            set
            {
                _selectedRamCatalog = value;
                OnPropertyChanged();
                // 選択されたデータの Format をキーに、マスターから検索して更新
                UpdateSelectedFormatDetail();
            }
        }

        private FormatData _selectedFormatDetail;
        public FormatData SelectedFormatDetail
        {
            get => _selectedFormatDetail;
            set
            {
                _selectedFormatDetail = value;
                OnPropertyChanged();
            }
        }

        private void UpdateSelectedFormatDetail()
        {
            if (_selectedRamCatalog == null || string.IsNullOrEmpty(_selectedRamCatalog.Format))
            {
                SelectedFormatDetail = null;
                return;
            }

            // FormatList(マスター)の中から、Idが一致するものを探す
            SelectedFormatDetail = FormatList.FirstOrDefault(f => f.Id == _selectedRamCatalog.Format);
        }

        private FormatData _selectedFormat;
        public ObservableCollection<FormatData> FormatList { get; } = new();
        public FormatData SelectedFormat
        {
            get => _selectedFormat;
            set
            {
                _selectedFormat = value;
                OnPropertyChanged();
            }
        }

        private RamData _selectedRamdata;
        public ObservableCollection<RamData> RamdataList { get; } = new();
        public RamData SelectedRamdata { 
            get => _selectedRamdata; 
            set{
                _selectedRamdata = value;
                OnPropertyChanged();
                // 全データの選択フラグを更新
                foreach (var item in RamdataList)
                {
                    item.IsSelected = (item == _selectedRamdata);
                }
            }
        }

        void UpdateSheet()
        {
            SelectedSheetData.Clear();

            if (_selectedSheet == null) return;

            foreach (var d in allSheets[_selectedSheet])
                SelectedSheetData.Add(d);
        }

        public void LoadExcel(string path)
        {
            SheetNames.Clear();
            SelectedSheetData.Clear();
            FormatList.Clear();

            allSheets = ExcelLoader.Load(path);
            
            foreach (var name in allSheets.Keys)
                SheetNames.Add(name);

            if (SheetNames.Count > 0)
                SelectedSheet = SheetNames[0];

            var formats = ExcelLoader.LoadFormats(path);
            foreach (var f in formats)
                FormatList.Add(f);

            if (SheetNames.Count > 0)
                SelectedSheet = SheetNames[0];

        }


        public event PropertyChangedEventHandler PropertyChanged;

        void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        public void ClearSelection()
        {
            SelectedRamdata = null; // Setter内のループで全IsSelectedがfalseになります
        }

        // 削除メソッド
        public void RemoveSelectedRam()
        {
            if (SelectedRamdata != null)
            {
                RamdataList.Remove(SelectedRamdata);
                ClearSelection();
            }
        }
    }
}
