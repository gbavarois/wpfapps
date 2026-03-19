using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;

namespace WpfApp1.Helpers
{
    public static class ColorHelper
    {
        public static string GetColorCode(string index) => index switch
        {
            "0" => "#000000",
            "1" => "#4040FF",
            "2" => "#FF4040",
            "3" => "#FF00FF",
            "4" => "#00FF00",
            "5" => "#00FFFF",
            "6" => "#FFFF00",
            "7" => "#FFFFFF",
            "8" => "#FF8000",
            "9" => "#80FF00",
            _ => "#000000"
        };

        public static Brush GetBrushFromIndex(string index)
        {
            var colorCode = GetColorCode(index);
            return (Brush)new BrushConverter().ConvertFromString(colorCode);
        }

        public static string GetColorIndex(Brush brush)
        {
            if (brush is not SolidColorBrush scb) return "0";

            string color = scb.Color.ToString().ToUpper();

            return color switch
            {
                "#FF000000" => "0",
                "#FF4040FF" => "1",
                "#FFFF4040" => "2",
                "#FFFF00FF" => "3",
                "#FF00FF00" => "4",
                "#FF00FFFF" => "5",
                "#FFFFFF00" => "6",
                "#FFFFFFFF" => "7",
                "#FFFF8000" => "8",
                "#FF80FF00" => "9",
                _ => "0"
            };
        }
    }
}
