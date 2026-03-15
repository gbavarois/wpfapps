using Microsoft.Win32;
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
using WpfApp1.ViewModels;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        // 0～9用のコマンドを定義
        public static RoutedCommand ColorCommand = new RoutedCommand();

        MainViewModel vm = new MainViewModel();

        public MainWindow()
        {
            InitializeComponent();

            DataContext = vm;

            // Ctrl + 0 ～ 9 のキー入力をコマンドに結びつける
            for (int i = 0; i <= 9; i++)
            {
                Key key = (Key)Enum.Parse(typeof(Key), "D" + i);
                CommandBindings.Add(new CommandBinding(ColorCommand, ColorCommand_Executed));
                InputBindings.Add(new KeyBinding(ColorCommand, key, ModifierKeys.Control) { CommandParameter = i.ToString() });
            }
        }

        private void LoadRamData_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog();

            dlg.Filter = "Excel|*.xlsx";

            if (dlg.ShowDialog() == true)
            {
                vm.LoadExcel(dlg.FileName);
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
            if (RamDataPane != null)
            {
                RamDataPane.IsVisible = true;
                RamDataPane.IsVisible = true;
                //TablePane.Show();
            }
        }

        private bool _isRamMoving = false;

        // Ctrlキー押下時のカーソル変更(矢印にする)
        private void MainEditor_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if ((e.Key == Key.LeftCtrl || e.Key == Key.RightCtrl) && !_isRamMoving)
            {
                Mouse.OverrideCursor = Cursors.Arrow;
            }
        }

        // Ctrlキーを離した時にカーソルを戻す
        private void MainEditor_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if ((e.Key == Key.LeftCtrl || e.Key == Key.RightCtrl) && !_isRamMoving)
            {
                Mouse.OverrideCursor = null;
            }
        }

        // フォーカスが外れた時にカーソルが固まるのを防ぐ
        private void MainEditor_LostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            _isRamMoving = false;
            Mouse.OverrideCursor = null;
        }

        // Ctrl + 左クリックでドラッグ開始
        private void MainEditor_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (Keyboard.Modifiers == ModifierKeys.Control)
            {
                _isRamMoving = true;
                MainEditor.CaretBrush = Brushes.Red;    // ドラッグ中はカレットを赤くする
                Mouse.OverrideCursor = Cursors.SizeAll; // カーソルを十字矢印に
                MainEditor.CaptureMouse();
                e.Handled = true;                       // テキスト選択を防ぐ
            }
        }

        // マウス移動に合わせてカレットを飛ばす
        private void MainEditor_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (_isRamMoving)
            {
                Point mousePos = e.GetPosition(MainEditor);
                // マウス座標に一番近い「文字の隙間」を取得
                TextPointer tp = MainEditor.GetPositionFromPoint(mousePos, snapToText: true);
                if (tp != null)
                {
                    MainEditor.CaretPosition = tp; // カレットを移動
                }
            }
        }

        // マウスを離した瞬間にRAM座標を確定
        private void MainEditor_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (_isRamMoving)
            {
                _isRamMoving = false;
                MainEditor.CaretBrush = Brushes.Lime; // カレットを戻す
                // マウスアップ時にCtrlがまだ押されていれば「矢印」に、そうでなければ「I」に戻す
                Mouse.OverrideCursor = Keyboard.Modifiers.HasFlag(ModifierKeys.Control) ? Cursors.Arrow : null;
                MainEditor.ReleaseMouseCapture();

                // 現在のカレット位置から「行・桁」を取得
                //GetCaretLineColumn(out int row, out int col);

                // ここでRAM情報のデータ(Row, Column)を更新し、
                // 背面(HighlightEditor)の特定の文字範囲を塗り直す処理を呼ぶ
                //UpdateHighlight(row, col);
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
