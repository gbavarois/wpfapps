using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;

namespace WpfApp1.Helpers
{
    public static class VisualTreeHelperExtensions
    {
        // メソッドにも static を付ける
        public static T? GetVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            if (parent == null) return null;
            // VisualTreeHelper.GetChildrenCount(...) クラス名から直接呼ぶ
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T t) return t;
                var result = GetVisualChild<T>(child);
                if (result != null) return result;
            }
            return null;
        }
    }
}
