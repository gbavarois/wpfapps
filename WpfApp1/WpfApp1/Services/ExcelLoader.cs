using ClosedXML.Excel;
using ExcelDataReader;
using System.Data;
using System.IO;
using WpfApp1.Models;

namespace WpfApp1.Services
{
    public static class ExcelLoader
    {
        public static Dictionary<string, List<RamCatalog>> Load(string path)
        {
            var result = new Dictionary<string, List<RamCatalog>>();

            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var wb = new XLWorkbook(stream);

            foreach (var ws in wb.Worksheets)
            {
                // シート名判定 (swrt_0 ～ swrt_15)
                if (!ws.Name.StartsWith("swrt_") ||
                    !int.TryParse(ws.Name.AsSpan(5), out int num) ||
                    num < 0 || num > 15) continue;

                var firstRow = ws.FirstRowUsed();

                // 各項目の列番号を動的に取得
                int cData = GetColNum(firstRow, "Data");
                int cSymbol = GetColNum(firstRow, "Symbol");
                int cType = GetColNum(firstRow, "Type");
                int cLen = GetColNum(firstRow, "Length");
                int cUnit = GetColNum(firstRow, "Unit");
                int cLsb = GetColNum(firstRow, "LSB");
                int cFmt = GetColNum(firstRow, "format");
                int cNote = GetColNum(firstRow, "note");
                int cAddress = GetColNum(firstRow, "address");

                var list = new List<RamCatalog>();

                // データ行（2行目以降）をスキャン
                foreach (var row in ws.RowsUsed().Skip(1))
                {
                    var vData = row.Cell(cData).GetString();
                    var vSymbol = row.Cell(cSymbol).GetString();
                    var vType = row.Cell(cType).GetString();
                    var vLen = row.Cell(cLen).GetString();
                    var vUnit = row.Cell(cUnit).GetString();
                    var vLsb = row.Cell(cLsb).GetString();
                    var vFmt = row.Cell(cFmt).GetString();
                    var vNote = row.Cell(cNote).GetString();
                    var vAddress = row.Cell(cAddress).GetString();

                    // 指定された全列が空白なら、そこでこのシートの読み込みを終了
                    if (string.IsNullOrWhiteSpace(vData) && string.IsNullOrWhiteSpace(vSymbol) &&
                        string.IsNullOrWhiteSpace(vType) && string.IsNullOrWhiteSpace(vLen) &&
                        string.IsNullOrWhiteSpace(vUnit) && string.IsNullOrWhiteSpace(vLsb) &&
                        string.IsNullOrWhiteSpace(vFmt) && string.IsNullOrWhiteSpace(vNote) &&
                        string.IsNullOrWhiteSpace(vAddress))
                    {
                        break;
                    }

                    list.Add(new RamCatalog
                    {
                        Data = vData,
                        Symbol = vSymbol,
                        Type = vType,
                        Length = vLen,
                        Unit = vUnit,
                        LSB = vLsb,
                        FormatId = vFmt,
                        Note = vNote,
                        Address = vAddress
                    });
                }
                result[ws.Name] = list;
            }
            return result;
        }
        private static int GetColNum(IXLRow headerRow, string columnName)
        {
            // 列名が見つからない場合は 0 を返し、row.Cell(0) は空のセルとして扱われるようにする
            return headerRow.Cells().FirstOrDefault(c => c.GetString() == columnName)?.Address.ColumnNumber ?? 0;
        }

        public static List<FormatData> LoadFormats(string path)
        {
            var list = new List<FormatData>();

            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var wb = new XLWorkbook(stream);
            var ws = wb.Worksheet("表記フォーマット");

            var firstRow = ws.FirstRowUsed();

            // 各項目の列番号を動的に取得
            int cId = GetColNum(firstRow, "id");
            int cCode = GetColNum(firstRow, "code");
            int cLength = GetColNum(firstRow, "length");
            int cPlaceholder = GetColNum(firstRow, "placeholder");
            int cDescription = GetColNum(firstRow, "description");

            // データ行（2行目以降）をスキャン
            foreach (var row in ws.RowsUsed().Skip(1))
            {
                var vId = row.Cell(cId).GetString();
                var vCode = row.Cell(cCode).GetString();
                var vLength = row.Cell(cLength).GetValue<int>();
                var vPlaceholder = row.Cell(cPlaceholder).GetString();
                var vDescription = row.Cell(cDescription).GetString();
                list.Add(new FormatData
                {
                    Id = vId,
                    Code = vCode,
                    Length = vLength,
                    Placeholder = vPlaceholder,
                    Description = vDescription
                });
            }

            return list;
        }
    }
}