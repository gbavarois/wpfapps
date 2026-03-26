using Microsoft.Win32;
using System;
using System.Globalization;
using System.Runtime.Intrinsics.Arm;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using WpfApp1.Helpers;
using WpfApp1.Models;
using WpfApp1.Services;
using WpfApp1.ViewModels;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private const double CharWidth = 7.0;
        private const double LineHeightValue = 14.0;

        // 現在開いているファイルのフルパスを保持する変数
        private string? _currentFilePath;

        // 0～9用のコマンドを定義
        public static RoutedCommand ColorCommand = new RoutedCommand();


        public MainWindow()
        {
            InitializeComponent();
            // ViewModelをセット
            this.DataContext = new MainViewModel();
            this.CommandBindings.Add(new CommandBinding(ColorCommand, ColorCommand_Executed));
        }

        // ドラッグ移動：ピクセル計算の結果をデータモデルに戻すだけ
        private void Thumb_DragDelta(object sender, DragDeltaEventArgs e)
        {
            if (sender is Thumb thumb && thumb.DataContext is RamItemViewModel data)
            {
                // Canvas上での現在の位置を取得
                var parent = VisualTreeHelper.GetParent(thumb) as ContentPresenter;
                double left = Canvas.GetLeft(parent) + e.HorizontalChange;
                double top = Canvas.GetTop(parent) + e.VerticalChange;

                // 文字単位の座標にスナップしてデータを更新
                data.Column = (int)Math.Max(0, Math.Round(left / CharWidth));
                data.Row = (int)Math.Max(0, Math.Round(top / LineHeightValue));
            }
        }
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

        // Canvasをクリックした際に選択解除
        private void MainEditor_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.OriginalSource is not Border && e.OriginalSource is not Shape)
            {
                (DataContext as MainViewModel)?.ClearSelectionCommand.Execute(null);
            }
        }

        // スクロール同期（これはViewの仕事として残す）
        private void MainEditor_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            // ScrollChangedの中などの記述を修正
            var canvas = MyItemsControl.Template.FindName("RamCanvas", MyItemsControl) as Canvas;
            if (canvas != null)
            {
                canvas.RenderTransform = new TranslateTransform(-e.HorizontalOffset, -e.VerticalOffset);
            }

        }

        private void Thumb_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is Thumb thumb && thumb.DataContext is RamItemViewModel data)
            {
                var vm = (MainViewModel)this.DataContext;
                vm.ActiveTab.SelectedRam = data;
            }
        }

        // メニューの「RAMデータ読込」クリック
        private void LoadRamData_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
            };

            if (dialog.ShowDialog() == true)
            {
                var vm = (MainViewModel)this.DataContext;
                // ViewModelのコマンドを、パスを引数にして実行する
                vm.LoadExcelCommand.Execute(dialog.FileName);
            }
        }

        private void NewFile_Click(object sender, RoutedEventArgs e)
        {
            MainEditor.Document = new FlowDocument();
            _currentFilePath = null;
        }

        // JSON読み込みメニュー用
        private void OpenFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
            };

            if (dialog.ShowDialog() == true)
            {
                var vm = (MainViewModel)this.DataContext;

                // Command.Execute(...) ではなく、直接メソッドを呼ぶ
                var saveData = vm.LoadFromJson(dialog.FileName);

                if (saveData != null)
                {
                    var service = new JsonEditorService();
                    service.RestoreText(MainEditor, saveData.Lines);
                    service.ApplyColorInfo(MainEditor, saveData.Colors);
                }

                _currentFilePath = dialog.FileName;
            }
        }
        private void SaveFile_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_currentFilePath))
            {
                SaveAsFile_Click(sender, e);
                return;
            }

            var service = new JsonEditorService();
            var vm = (MainViewModel)DataContext;
            var data = service.CreateSaveData(MainEditor, vm.RamdataList);

            service.SaveToJson(data, _currentFilePath);
        }
        private void SaveAsFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
            };

            if (dialog.ShowDialog() == true)
            {
                var service = new JsonEditorService();
                var vm = (MainViewModel)DataContext;
                var data = service.CreateSaveData(MainEditor, vm.RamdataList);

                service.SaveToJson(data, dialog.FileName);

                _currentFilePath = dialog.FileName;
            }
        }

        private void SetColor_Click(object sender, RoutedEventArgs e)
        {
            // メニュー項目またはショートカットから Tag（色コード）を取得
            string colorCode = (sender as MenuItem)?.Tag?.ToString() ?? "0";
            ApplyColorToSelection(colorCode);
        }

        private void ColorCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            if (e.Parameter is string index)
            {
                ApplyColorToSelection(index);
            }
        }

        // 実際の着色ロジック
        private void ApplyColorToSelection(string colorIndex)
        {
            var range = MainEditor.Selection;
            if (!range.IsEmpty && !range.IsEmpty)
            {
                // ColorHelperを使ってインデックスからブラシを取得
                var brush = ColorHelper.GetBrushFromIndex(colorIndex);
                range.ApplyPropertyValue(TextElement.ForegroundProperty, brush);
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

            var vm = (MainViewModel)DataContext;
            vm.CurrentRow = lineCount;
            vm.CurrentColumn = columnCount;

            if (StatusLineColumn != null)
            {
                StatusLineColumn.Text = $"行:{lineCount}, 桁:{columnCount}";
            }
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

    }

}
