using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.ComponentModel;
using WpfApp1.Models;

namespace WpfApp1.ViewModels
{
    public partial class RamItemViewModel : ObservableObject
    {
        private readonly RamLayout _model;      // 保存対象のピュアなデータ
        private readonly MainViewModel _main;   // マスター参照用

        public RamItemViewModel(RamLayout model, MainViewModel main)
        {
            _model = model;
            _main = main;
        }

        // 保存データ（Model）をそのまま取得したい時用
        public RamLayout Model => _model;

        // Modelの値をプロパティとして公開（変更されたらModelも書き換える）
        public int Row
        {
            get => _model.Row;
            set { if (SetProperty(_model.Row, value, _model, (m, v) => m.Row = v)) OnPropertyChanged(); }
        }

        public int Column
        {
            get => _model.Column;
            set { if (SetProperty(_model.Column, value, _model, (m, v) => m.Column = v)) OnPropertyChanged(); }
        }

        public string FormatId
        {
            get => _model.FormatId;
            set
            {
                if (SetProperty(_model.FormatId, value, _model, (m, v) => m.FormatId = v))
                {
                    OnPropertyChanged();
                    // FormatIdが変わると、以下のプロパティの結果も変わるため通知する
                    OnPropertyChanged(nameof(Format));
                    OnPropertyChanged(nameof(Length));
                    OnPropertyChanged(nameof(Placeholder));
                    OnPropertyChanged(nameof(IsValid));
                }
            }
        }

        public string Symbol => _model.Symbol;

        // --- 参照・計算プロパティ（Modelは持たず、ViewModelが解決する） ---

        public RamCatalog? Catalog => _main.CurrentCatalogs.FirstOrDefault(c => c.Symbol == _model.Symbol);

        public FormatData? Format => _main.FormatList.FirstOrDefault(f => f.Id == FormatId);

        public string Placeholder => Format?.Placeholder ?? string.Empty;

        public bool IsValid => Catalog != null && Format != null;



        public string Data => Catalog?.Data ?? "(不明)";

        public int Length => Format?.Length ?? 8;

        public string ComputedAddress
        {
            get
            {
                if (int.TryParse(Catalog?.Address, out var baseAddr))
                    return (baseAddr + _model.Offset).ToString();
                return Catalog?.Address ?? "";
            }
        }
    }
}
