using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WpfApp1.Services;

namespace WpfApp1.Models
{
    public class RamTableMaster
    {
        private Dictionary<string, List<RamCatalog>> _pages = new();
        public IReadOnlyDictionary<string, List<RamCatalog>> Pages => _pages;

        private List<FormatData> _formats = new();
        public IReadOnlyList<FormatData> Formats => _formats;

        public void Load(string filePath)
        {
            _pages = ExcelLoader.LoadCatalogs(filePath);
            _formats = ExcelLoader.LoadFormats(filePath);

            _pages.ElementAt(0).Value.Add(new RamCatalog
            {
                Data = "時刻データ",
                Symbol = "-",
                Type = "1",
                Length = "1",
                Unit = "-",
                LSB = "-",
                FormatId = "TIME",
                Note = "右下隅に配置する時刻データ",
                Address = "0F8"
            });
            _formats.Add(new FormatData
                {
                    Id = "TIME",
                    Code = "-1",
                    Length = 11,
                    Placeholder = "HH:MM:SS.00",
                    Description = "時刻データ"
                });
        }

        //private Dictionary<string, List<RamCatalog>> _pages = new();
        //private List<FormatData> _formats = new(); // 内部保持

        //public List<string> SheetNames => _pages.Keys.ToList();

        //// ViewModelから参照できるように公開
        //public List<FormatData> Formats => _formats;

        //public void Load(string filePath)
        //{
        //    _pages = ExcelLoader.LoadCatalogs(filePath);
        //    _formats = ExcelLoader.LoadFormats(filePath);
        //}

        //public List<RamCatalog> GetCatalogsBySheet(string sheetName)
        //{
        //    return _pages.TryGetValue(sheetName, out var list) ? list : new List<RamCatalog>();
        //}

        public RamCatalog? FindCatalogBySymbol(string symbol)
        {
            // 全てのシート(Value)を巡回して、一致する Symbol を探す
            return _pages.Values
                         .SelectMany(list => list)
                         .FirstOrDefault(c => c.Symbol == symbol);
        }
    }
}
