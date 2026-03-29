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
        private List<FormatData> _formats = new(); // 内部保持

        public List<string> SheetNames => _pages.Keys.ToList();

        // ViewModelから参照できるように公開
        public List<FormatData> Formats => _formats;

        public void Load(string filePath)
        {
            _pages = ExcelLoader.LoadCatalogs(filePath);
            _formats = ExcelLoader.LoadFormats(filePath);
        }

        public List<RamCatalog> GetCatalogsBySheet(string sheetName)
        {
            return _pages.TryGetValue(sheetName, out var list) ? list : new List<RamCatalog>();
        }

        public RamCatalog? FindCatalogBySymbol(string symbol)
        {
            // 全てのシート(Value)を巡回して、一致する Symbol を探す
            return _pages.Values
                         .SelectMany(list => list)
                         .FirstOrDefault(c => c.Symbol == symbol);
        }
    }
}
