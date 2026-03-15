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

        public ObservableCollection<string> SheetNames { get; } = new();
        public ObservableCollection<RamData> SelectedSheetData { get; } = new();
        
        public RamData SelectedRam { get; set; }

        Dictionary<string, List<RamData>> allSheets = new();

        string selectedSheet;

        public string SelectedSheet
        {
            get => selectedSheet;
            set
            {
                selectedSheet = value;
                OnPropertyChanged();
                UpdateSheet();
            }
        }

        void UpdateSheet()
        {
            SelectedSheetData.Clear();

            if (selectedSheet == null) return;

            foreach (var d in allSheets[selectedSheet])
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
    }
}
