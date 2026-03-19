using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Spreadsheet;
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
        private string _formatId;
        private FormatData _format;
        private int _offset;
        private string _symbol;
        private bool _isSelected;

        public IEnumerable<FormatData> FormatSource { get; set; }

        // 元データのカタログ情報（単位やアドレスなどのマスター参照）
        public RamCatalog Catalog { get; set; }

        public int Row
        {
            get => _row;
            set { _row = value; OnPropertyChanged(); }
        }
        public int Column
        {
            get => _column;
            set { _column = value; OnPropertyChanged(); }
        }
        public string FormatId
        {
            get => _formatId;
            set
            {
                if (_formatId == value) return;

                _formatId = value;
                OnPropertyChanged();

                if (FormatSource != null)
                {
                    Format = FormatSource.FirstOrDefault(f => f.Id == _formatId);
                }
                else
                {
                    Format = null;
                }

                OnPropertyChanged(nameof(IsValid));
            }
        }
        public FormatData Format
        {
            get => _format;
            set
            {
                if (_format == value) return;

                _format = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(Length));
                OnPropertyChanged(nameof(Placeholder));
            }
        }
        public int Offset
        {
            get => _offset;
            set
            {
                _offset = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ComputedAddress));
            }
        }

        public string Symbol
        {
            get => Catalog?.Symbol ?? _symbol ?? "(不明)";
            set
            {
                _symbol = value;
                OnPropertyChanged();
            }
        }

        public bool IsSelected
        {
            get => _isSelected;
            set { _isSelected = value; OnPropertyChanged(); }
        }

        // 表示用のプロパティ（カタログから引っ張る）
        public string Data => Catalog?.Data ?? "(不明)";
        public string BaseAddress => Catalog?.Address;
        public string ComputedAddress
        {
            get
            {
                if (int.TryParse(Catalog?.Address, out var baseAddr))
                    return (baseAddr + Offset).ToString();
                return Catalog?.Address;
            }
        }
        public int Length => Format?.Length ?? 0;
        public string Placeholder => Format?.Placeholder;
        public bool IsValid =>  Catalog != null &&
                                (FormatSource == null || Format != null);

        public void ResolveFormat(IEnumerable<FormatData> formatList)
        {
            //Format = formatList.FirstOrDefault(f => f.Id == FormatId);
            Format = formatList.FirstOrDefault(f => string.Equals(f.Id?.Trim(), FormatId?.Trim(), StringComparison.OrdinalIgnoreCase));
            OnPropertyChanged(nameof(Format));
            OnPropertyChanged(nameof(IsValid));
            OnPropertyChanged(nameof(Length));
            OnPropertyChanged(nameof(Placeholder));
        }

        public void NotifyCatalogChanged()
        {
            OnPropertyChanged(nameof(Catalog));
            OnPropertyChanged(nameof(Symbol));
            OnPropertyChanged(nameof(Data));
            OnPropertyChanged(nameof(Address));
            OnPropertyChanged(nameof(IsValid));
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
