using System;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        // 0～9用のコマンドを定義
        public static RoutedCommand ColorCommand = new RoutedCommand();

        public MainWindow()
        {
            InitializeComponent();

            // Ctrl + 0 ～ 9 のキー入力をコマンドに結びつける
            for (int i = 0; i <= 9; i++)
            {
                Key key = (Key)Enum.Parse(typeof(Key), "D" + i);
                CommandBindings.Add(new CommandBinding(ColorCommand, ColorCommand_Executed));
                InputBindings.Add(new KeyBinding(ColorCommand, key, ModifierKeys.Control) { CommandParameter = i.ToString() });
            }
        }

        // コマンド実行時の共通処理
        private void ColorCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            // 実行パラメータ（0～9）から色を決定
            string index = e.Parameter as string;
            // 先ほど作った共通メソッドを呼び出す（擬似的にMenuItemを偽装して再利用）
            var dummy = new MenuItem { Tag = GetColorCode(index) };
            ApplyColor(dummy.Tag.ToString());
        }

        private void SetColor_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem != null) ApplyColor(menuItem.Tag.ToString());
        }

        // 共通の色変更ロジック
        private void ApplyColor(string colorCode)
        {
            if (MainEditor.Selection.IsEmpty) return;
            var color = (Color)ColorConverter.ConvertFromString(colorCode);
            MainEditor.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, new SolidColorBrush(color));
        }

        // 数字からカラーコードを引くヘルパー
        private string GetColorCode(string index) => index switch
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
            _ => "#FFFFFF"
        };

        // --- メニュー項目のクリック処理 ---

        // 「表記フォーマット パネル」を表示する
        private void ShowFormatPane_Click(object sender, RoutedEventArgs e)
        {
            // AvalonDockの機能で、非表示状態から再表示させます
            if (FormatPane != null)
            {
                FormatPane.IsVisible = true;
                //FormatPane.Show();
                FormatPane.IsSelected = true;
            }
        }

        // 「テーブルデータ パネル」を表示する
        private void ShowTablePane_Click(object sender, RoutedEventArgs e)
        {
            if (TablePane != null)
            {
                TablePane.IsVisible = true;
                TablePane.IsVisible = true;
                //TablePane.Show();
            }
        }

        private void MainEditor_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            // Ctrl と Shift が両方押されているか
            var modifiers = Keyboard.Modifiers;
            if ((modifiers & ModifierKeys.Control) != 0 && (modifiers & ModifierKeys.Shift) != 0)
            {
                // F1 (Key.F1) ～ F10 (Key.F10) かを判定
                if (e.Key >= Key.F1 && e.Key <= Key.F10)
                {
                    // F1なら "0", F2なら "1" ... F10なら "9" として色を決定
                    int index = e.Key - Key.F1;
                    ApplyColor(GetColorCode(index.ToString()));

                    e.Handled = true; // RichTextBox側の標準動作をさせない
                }
            }
        }

        // --- リッチテキストボックスのカーソル位置計算 ---
        private void MainEditor_SelectionChanged(object sender, RoutedEventArgs e)
        {
            // カーソルの現在位置を取得
            TextPointer caretPos = MainEditor.CaretPosition;

            // 行（Line）の計算
            int lineCount = 0;
            // 現在の行の先頭位置を取得
            TextPointer lineStart = caretPos.GetLineStartPosition(0);

            // ドキュメントの最初から、現在の行まで何回改行があるかカウント
            TextPointer temp = lineStart;
            while (temp != null && temp.CompareTo(MainEditor.Document.ContentStart) > 0)
            {
                // 1行上に移動を試みる
                TextPointer prevLine = temp.GetLineStartPosition(-1);
                if (prevLine != null && prevLine.CompareTo(temp) < 0)
                {
                    temp = prevLine;
                    lineCount++;
                }
                else
                {
                    break;
                }
            }

            // 文字（列）の計算：行の先頭からカーソル位置までの文字数
            int columnCount = 0;
            if (lineStart != null)
            {
                columnCount = lineStart.GetOffsetToPosition(caretPos) - 1;
            }

            // ステータスバーのテキストを更新（行、文字ともに0から開始）
            if (StatusLineColumn != null)
            {
                StatusLineColumn.Text = $"行:{lineCount}, 文字:{columnCount}";
            }
        }
    }
}
