using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
using WpfApp1;
using WpfApp1.Helpers;
using WpfApp1.Models;
using WpfApp1.Services;
using WpfApp1.ViewModels;

namespace WpfApp1.Views
{
    /// <summary>
    /// UserControl1.xaml の相互作用ロジック
    /// </summary>
    public partial class EditorView : UserControl
    {
        private const double CharWidth = 7.0;
        private const double LineHeightValue = 14.0;

        public EditorView()
        {
            InitializeComponent();
        }

        // Canvasをクリックした際に選択解除
        private void MainEditor_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            // リッチテキストの「何もないところ」や「文字」をクリックした時
            if (this.DataContext is DisplayEditorViewModel tabVM)
            {
                // 現在選択されている RAM を解除（nullを代入）
                tabVM.SelectedRam = null;
            }
        }

        // スクロール同期（これはViewの仕事として残す）
        private void MainEditor_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            var rtb = sender as RichTextBox;
            if (rtb == null) return;

            // 1. RichTextBox と同じ階層（Grid）にある ItemsControl を探す
            // 親の Grid を取得
            var parentGrid = rtb.Parent as Grid;
            if (parentGrid == null) return;

            // 2. Grid の子要素から ItemsControl (MyItemsControl) を探す
            var itemsControl = parentGrid.Children.OfType<ItemsControl>().FirstOrDefault(x => x.Name == "MyItemsControl");
            if (itemsControl == null) return;

            // 3. ItemsControl の中にある Canvas (RamCanvas) を探す
            // ItemsControl がロード完了していないと Template.FindName は null になるので注意
            if (itemsControl.Template != null)
            {
                var canvas = VisualTreeHelperExtensions.GetVisualChild<Canvas>(itemsControl);
                if (canvas != null)
                {
                    // スクロール量を Canvas のズレとして適用（同期）
                    canvas.RenderTransform = new TranslateTransform(-e.HorizontalOffset, -e.VerticalOffset);
                }
            }
        }

        // --- リッチテキストボックスのカーソル位置計算 ---
        private void MainEditor_SelectionChanged(object sender, RoutedEventArgs e)
        {
            if (sender is not RichTextBox rtb) return;

            TextPointer caretPos = rtb.CaretPosition;
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

            if (this.DataContext is DisplayEditorViewModel tabVM)
            {
                tabVM.CurrentRow = lineCount;
                tabVM.CurrentColumn = columnCount;
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

        private void SetColor_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem menuItem && menuItem.Tag is string colorIndex)
            {
                ApplyColorToSelection(colorIndex);
            }
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

        private void Thumb_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is Thumb thumb)
            {
                // 1. フォーカスを強制的に Thumb (またはその親) に移す
                // これにより、Deleteキーなどのコマンドがこのコンテキストで有効になります
                thumb.Focus();

                if (thumb.DataContext is RamItemViewModel data && this.DataContext is DisplayEditorViewModel tabVM)
                {
                    tabVM.SelectedRam = data;
                    // 右クリックメニューを出すためにイベントを「処理済み」にしない
                }
            }
        }

        // 実際の着色ロジック
        public void ApplyColorToSelection(string colorIndex)
        {
            var rtb = this.MainEditor;
            if (rtb == null) return;

            var range = rtb.Selection;
            if (!range.IsEmpty && !range.IsEmpty)
            {
                // ColorHelperを使ってインデックスからブラシを取得
                var brush = ColorHelper.GetBrushFromIndex(colorIndex);
                range.ApplyPropertyValue(TextElement.ForegroundProperty, brush);
            }
        }

        public EditorData GetEditorData()
        {
            var service = new JsonEditorService();
            var vm = (DisplayEditorViewModel)this.DataContext;
            return service.CreateSaveEditorData(vm.DisplayName, this.MainEditor, vm.PlacedRams.Select(r => r.Model));
        }

        private void EditorView_Loaded(object sender, RoutedEventArgs e)
        {
            if (this.DataContext is DisplayEditorViewModel tabVM && tabVM.RestoreData != null)
            {
                var service = new JsonEditorService();
                var data = tabVM.RestoreData;

                this.MainEditor.Document.Blocks.Clear();
                service.RestoreText(this.MainEditor, data.Lines);
                service.RestoreColors(this.MainEditor, data.Colors);

                // 復元が終わったらメモリ解放のために消しておく
                tabVM.RestoreData = null;
            }
        }
    }
}
