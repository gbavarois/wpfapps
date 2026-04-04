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
    public partial class DisplayEditorViewModel : ObservableObject
    {
        private readonly MainViewModel _main;

        [ObservableProperty] private int _currentRow;
        [ObservableProperty] private int _currentColumn;

        [ObservableProperty] private string _displayName = "新規ディスプレイ";

		public event Action<string>? ApplyColorRequested;

		// このディスプレイに配置されているRAMのリスト
		public ObservableCollection<RamItemViewModel> PlacedRams { get; } = new();

        // このディスプレイで選択されているRAM
        [ObservableProperty] private RamItemViewModel? _selectedRam;

        public DisplayEditorViewModel(MainViewModel main)
        {
            _main = main;
        }
        
        public MainViewModel Main => _main;

        public EditorData? RestoreData { get; set; }

		[RelayCommand]
		private void ApplyColor(string colorIndex)
		{
			ApplyColorRequested?.Invoke(colorIndex);
		}

        public EditorData GetSaveData()
        {
            return new EditorData
            {
                Title = DisplayName,
                Rams = PlacedRams.Select(r => r.ToModel()).ToList()
            };
        }

        partial void OnSelectedRamChanged(RamItemViewModel? value)
        {
            // 全部解除
            foreach (var ram in PlacedRams)
            {
                ram.IsSelected = false;
            }

            // 選択されたものだけON
            if (value != null)
            {
                value.IsSelected = true;
            }

            _main.RemoveRamCommand.NotifyCanExecuteChanged();
        }
    }
}
