using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WpfApp1.Models;

namespace WpfApp1.ViewModels
{
    public partial class EditorDisplayViewModel : ObservableObject
    {
        private readonly MainViewModel _main;

        [ObservableProperty] private string _displayNumber = "新規ディスプレイ";

        // このディスプレイに配置されているRAMのリスト
        public ObservableCollection<RamItemViewModel> PlacedRams { get; } = new();

        // このディスプレイで選択されているRAM
        [ObservableProperty] private RamItemViewModel? _selectedRam;

        public EditorDisplayViewModel(MainViewModel main)
        {
            _main = main;
        }

        [RelayCommand]
        public void AddRam(RamCatalog sourceCatalog)
        {
            // 1. Modelを作る（位置はとりあえず0,0など）
            var newModel = new RamLayout
            {
                Symbol = sourceCatalog.Symbol,
                FormatId = sourceCatalog.FormatId
            };

            // 2. ViewModelでラップしてリストに追加
            var vm = new RamItemViewModel(newModel, _main);
            PlacedRams.Add(vm);

            // 3. 選択状態にする
            SelectedRam = vm;
        }

        // 保存用データ（DTO）への変換
        public EditorSaveData CreateSaveData(System.Collections.Generic.List<string> lines, System.Collections.Generic.List<TextColorInfo> colors)
        {
            var data = new EditorSaveData { Lines = lines, Colors = colors };
            foreach (var ramVM in PlacedRams)
            {
                data.Rams.Add(ramVM.Model);
            }
            return data;
        }
    }
}
