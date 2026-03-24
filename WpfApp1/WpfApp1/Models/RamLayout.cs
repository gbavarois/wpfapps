using CommunityToolkit.Mvvm.ComponentModel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Collections.Generic;
using System.Linq;

namespace WpfApp1.Models
{
    public partial class RamLayout : ObservableObject
    {
        [ObservableProperty] private int _row;
        [ObservableProperty] private int _column;
        [ObservableProperty] private int _offset;
        [ObservableProperty] private bool _isSelected;

        [ObservableProperty] private IEnumerable<FormatData>? _formatSource;

        // IDが変わったら、Format実体を解決する
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(Length), nameof(Placeholder), nameof(IsValid))]
        private string _formatId = string.Empty;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(Length), nameof(Placeholder), nameof(IsValid))]
        private FormatData? _format;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(Symbol), nameof(Data), nameof(ComputedAddress), nameof(IsValid))]
        private RamCatalog? _catalog;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(Data))] // Symbolが変わったらData（表示用）も通知
        private string _symbol = string.Empty;

        // 表示用計算プロパティ
        
        public string Data => Catalog?.Data ?? "(不明)";
        public int Length => Format?.Length ?? 8;
        public string Placeholder => Format?.Placeholder ?? "";

        public string ComputedAddress
        {
            get
            {
                if (int.TryParse(Catalog?.Address, out var baseAddr))
                    return (baseAddr + Offset).ToString();
                return Catalog?.Address ?? "";
            }
        }

        public bool IsValid => Catalog != null && Format != null;

        // FormatIdが変更されたときに呼ばれるフック（自動でFormat実体を探す）
        public void ResolveReferences(IEnumerable<FormatData> formats, IEnumerable<RamCatalog> catalogs)
        {
            Format = formats.FirstOrDefault(f => f.Id?.Trim() == FormatId?.Trim());
            if (Catalog == null)
            {
                Catalog = catalogs.FirstOrDefault(c => c.Symbol == Symbol);
            }
        }
    }
}
