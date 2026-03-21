using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace WpfApp1
{
    public class PosConverter : IValueConverter
    {
        public double Scale { get; set; } // 8.4 や 14 を入れる
        public object Convert(object v, Type t, object p, CultureInfo c) => (int)v * Scale;
        public object ConvertBack(object v, Type t, object p, CultureInfo c) => throw new NotImplementedException();
    }
}
