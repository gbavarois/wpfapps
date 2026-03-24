using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.ComponentModel;
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

        // 保存用データ（DTO）への変換
        public EditorSaveData CreateSaveData(System.Collections.Generic.List<string> lines, System.Collections.Generic.List<TextColorInfo> colors)
        {
            var data = new EditorSaveData { Lines = lines, Colors = colors };
            foreach (var ramVM in PlacedRams)
            {
                data.Rams.Add(ramVM.GetModel());
            }
            return data;
        }
    }
}
