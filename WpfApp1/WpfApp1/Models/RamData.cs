using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace WpfApp1.Models
{
    public class RamData : INotifyPropertyChanged
    {
        private int _row;
        private int _column;
        private string _format;
        private bool _isSelected;

        // 元データのカタログ情報（単位やアドレスなどのマスター参照）
        public RamCatalog Catalog { get; set; }

        public int Row { get => _row; set { _row = value; OnPropertyChanged(); } }
        public int Column { get => _column; set { _column = value; OnPropertyChanged(); } }

        // 配置後に上書き（変更）可能なフォーマットID
        public string Format { get => _format; set { _format = value; OnPropertyChanged(); } }
        public bool IsSelected { get => _isSelected; set { _isSelected = value; OnPropertyChanged(); } }

        // 表示用のプロパティ（カタログから引っ張る）
        public string Symbol => Catalog?.Symbol;
        public string Address => Catalog?.Address;
        public string Data => Catalog?.Data;

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
