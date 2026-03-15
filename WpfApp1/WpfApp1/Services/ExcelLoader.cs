using ExcelDataReader;
using System.Data;
using System.IO;
using WpfApp1.Models;

namespace WpfApp1.Services
{
    public static class ExcelLoader
    {
        public static Dictionary<string, List<RamData>> Load(string path)
        {
            var result = new Dictionary<string, List<RamData>>();

            System.Text.Encoding.RegisterProvider(
                System.Text.CodePagesEncodingProvider.Instance);

            using var stream = File.Open(path, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);

            var ds = reader.AsDataSet();

            foreach (DataTable table in ds.Tables)
            {
                if (!table.TableName.StartsWith("swrt_"))
                    continue;

                if (!int.TryParse(table.TableName.Substring(5), out int num))
                    continue;

                if (num < 0 || num > 15)
                    continue;

                int colAddr = -1;
                int colCode = -1;

                // ヘッダ解析
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    var name = table.Rows[0][c]?.ToString();

                    if (name == "address") colAddr = c;
                    if (name == "code") colCode = c;
                }

                if (colAddr < 0 || colCode < 0)
                    continue;

                var list = new List<RamData>();

                // データ読み込み
                for (int r = 1; r < table.Rows.Count; r++)
                {
                    var addr = table.Rows[r][colAddr]?.ToString();

                    if (string.IsNullOrEmpty(addr))
                        break;

                    list.Add(new RamData
                    {
                        Row = r,
                        Column = 0,
                        Address = addr,
                        Format = table.Rows[r][colCode]?.ToString()
                    });
                }

                result[table.TableName] = list;
            }

            return result;
        }
    }
}