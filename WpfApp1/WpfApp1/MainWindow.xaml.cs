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
            if (RamCatalogPane != null)
            {
                RamCatalogPane.IsVisible = true;
                RamCatalogPane.IsVisible = true;
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
            TextPointer caretPos = MainEditor.CaretPosition;
            if (caretPos == null) return;

            // 1. 行（Line）の計算
            int lineCount = 0;
            TextPointer currentLineStart = caretPos.GetLineStartPosition(0);

            if (currentLineStart != null)
            {
                TextPointer runner = currentLineStart;
                while (true)
                {
                    // 1行上の先頭を取得（-1を指定）
                    TextPointer previousLineStart = runner.GetLineStartPosition(-1);

                    // 上の行が存在し、かつ今の行より前にある場合のみカウントアップ
                    if (previousLineStart != null && previousLineStart.CompareTo(runner) < 0)
                    {
                        lineCount++;
                        runner = previousLineStart;
                    }
                    else
                    {
                        // これ以上上の行がない（ドキュメントの先頭）
                        break;
                    }
                }
            }

            // 2. 桁（Column）の計算（全角2、半角1でカウント）
            int columnCount = 0;
            if (currentLineStart != null)
            {
                // 現在の行の先頭からカーソルまでのテキストを取得
                string textAtLine = new TextRange(currentLineStart, caretPos).Text;

                // 改行コードを除去してカウント
                foreach (char c in textAtLine)
                {
                    if (c == '\r' || c == '\n') continue;
                    columnCount += GetWidth(c);
                }
            }

            if (StatusLineColumn != null)
            {
                StatusLineColumn.Text = $"行:{lineCount}, 桁:{columnCount}";
            }
        }

        private string _lastText = string.Empty;

        private void MainEditor_TextChanged(object sender, TextChangedEventArgs e)
        {
            // 現在のテキストを取得（タグを除いた純粋な文字）
            string currentText = new TextRange(MainEditor.Document.ContentStart, MainEditor.Document.ContentEnd).Text;

            // 文字の中身が変わっていない（色指定などの書式変更だけ）なら何もしない
            if (currentText == _lastText) return;

            // 文字が変わった（入力された）場合は、即座に背景を更新
            _lastText = currentText;
            //UpdateHighlight();
        }

        

        private int GetWidth(char c)
        {
            // 半角カタカナの範囲 (U+FF61 ～ U+FF9F) は幅1として扱う
            if (c >= '\uFF61' && c <= '\uFF9F')
            {
                return 1;
            }

            // それ以外の 0xFF より大きい文字（漢字・ひらがな等）は幅2
            return c > 0xFF ? 2 : 1;
        }

        


        private void SetValidationMode(bool isEnabled)
        {
            if (isEnabled)
            {
                // 検証モード：前面を半透明にして、背面の黒との重なりを見せる
                MainEditor.Opacity = 0.6;
                MainEditor.Background = new SolidColorBrush(Color.FromArgb(50, 0, 0, 255)); // 薄い青
            }
            else
            {
                // 通常モード：前面を完全に透過させ、入力に集中させる
                MainEditor.Opacity = 1.0;
                MainEditor.Background = Brushes.Transparent;
            }
        }


        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            
        }

        // 基準となる定数
        private const double CharWidth = 7.0;       // 半角1文字の幅
        private const double LineHeightValue = 15.0; // 1行の高さ

        // メニューの「RAM追加」などから呼ぶテスト用メソッド
        private void AddTestRam_Click(object sender, RoutedEventArgs e)
        {
            // 例：現在のカーソル位置または指定位置に追加
            int testRow = 2;
            int testCol = 5;
            int testLen = 8; // 8桁分

            CreateRamVisual(testRow, testCol, testLen);
        }

        private void CreateRamVisual(int row, int col, int length)
        {
            System.Windows.Controls.Primitives.Thumb thumb = new System.Windows.Controls.Primitives.Thumb
            {
                Width = CharWidth * length,
                Height = LineHeightValue,
                Cursor = Cursors.SizeAll,
                Template = CreateRamTemplate()
            };

            // ドラッグ移動（スナップ付き）
            thumb.DragDelta += (s, e) =>
            {
                double left = Canvas.GetLeft(thumb) + e.HorizontalChange;
                double top = Canvas.GetTop(thumb) + e.VerticalChange;

                Canvas.SetLeft(thumb, Math.Max(0, Math.Round(left / CharWidth) * CharWidth));
                Canvas.SetTop(thumb, Math.Max(0, Math.Round(top / LineHeightValue) * LineHeightValue));
            };

            // 座標設定
            Canvas.SetLeft(thumb, col * CharWidth);
            Canvas.SetTop(thumb, row * LineHeightValue);

            RamCanvas.Children.Add(thumb);
        }

        private ControlTemplate CreateRamTemplate()
        {
            string xaml = @"
        <ControlTemplate xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'>
            <Border Background='#44FFFF00' />
        </ControlTemplate>";
            return (ControlTemplate)System.Windows.Markup.XamlReader.Parse(xaml);
        }
    }
}
