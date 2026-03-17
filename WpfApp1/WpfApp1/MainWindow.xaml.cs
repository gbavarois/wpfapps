using Microsoft.Win32;
using System;
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
using WpfApp1.Models;
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

            // RichTextBox 内の ScrollViewer から発生する ScrollChanged イベントを捕捉する
            MainEditor.AddHandler(ScrollViewer.ScrollChangedEvent, new ScrollChangedEventHandler(MainEditor_ScrollChanged));

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
        private void AddRamFromSelectedCatalog_Click(object sender, RoutedEventArgs e)
        {
            // DataContextからMainViewModelを取得
            var vm = this.DataContext as WpfApp1.ViewModels.MainViewModel;
            if (vm == null || vm.SelectedRamCatalog == null)
            {
                MessageBox.Show("左パネルの『RAMデータ一覧』から配置したいデータを選択してください。");
                return;
            }

            // 現在のテキストエディタのカーソル位置（行・桁）を取得
            TextPointer tp = MainEditor.CaretPosition;

            // 行の計算
            int row = 0;
            Paragraph currentPara = tp.Paragraph;
            foreach (var block in MainEditor.Document.Blocks)
            {
                if (block == currentPara) break;
                row++;
            }

            // 桁（列）の計算：段落の先頭からのオフセット
            TextRange range = new TextRange(currentPara.ContentStart, tp);
            int col = range.Text.Length;

            // 1. 配置データの実体 (RamData) を作成
            var newRam = new WpfApp1.Models.RamData
            {
                Catalog = vm.SelectedRamCatalog,     // マスター情報を紐付け
                Format = vm.SelectedRamCatalog.Format, // 初期値はカタログからコピー
                Row = row,
                Column = col
            };

            // 2. ViewModelのリストに追加
            // これにより右パネルの DataGrid (RamDataList) に一行追加される
            vm.RamdataList.Add(newRam);

            // 3. 配置したRAMのフォーマットから「桁数」を取得（なければデフォルト8）
            int length = 8;
            var formatInfo = vm.FormatList.FirstOrDefault(f => f.Id == newRam.Format);
            if (formatInfo != null)
            {
                length = formatInfo.Length;
            }

            // 4. Canvas上に四角（ハイライト）を生成して紐付け
            CreateRamVisual(newRam, length);
        }

        private void CreateRamVisual(RamData data, int length)
        {
            // 四角の幅を、カタログに紐付くフォーマットの Length から計算

            Thumb thumb = new Thumb
            {
                Width = CharWidth * length,
                Height = LineHeightValue,
                Template = CreateRamTemplate(),
                DataContext = data, // データモデルを直接紐付ける
                Tag = data
            };

            var vm = (MainViewModel)this.DataContext;

            // リストから自分が消されたら、Canvasからも消えるように監視
            vm.RamdataList.CollectionChanged += (s, e) =>
            {
                if (e.Action == System.Collections.Specialized.NotifyCollectionChangedAction.Remove ||
                    e.Action == System.Collections.Specialized.NotifyCollectionChangedAction.Reset)
                {
                    if (!vm.RamdataList.Contains(data))
                    {
                        RamCanvas.Children.Remove(thumb);
                    }
                }
            };

            // --- 以前の PropertyChanged 監視 (Widthや位置の更新) ---
            data.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(RamData.Format))
                {
                    var f = vm.FormatList.FirstOrDefault(x => x.Id == data.Format);
                    if (f != null) thumb.Width = CharWidth * f.Length;
                }
                else if (e.PropertyName == nameof(RamData.Row) || e.PropertyName == nameof(RamData.Column))
                {
                    UpdateThumbPos(thumb, data);
                }
            };

            // ドラッグ・クリックイベント（既存）
            thumb.DragDelta += (s, e) => {
                double left = Canvas.GetLeft(thumb) + e.HorizontalChange;
                double top = Canvas.GetTop(thumb) + e.VerticalChange;
                data.Column = (int)Math.Max(0, Math.Round(left / CharWidth));
                data.Row = (int)Math.Max(0, Math.Round(top / LineHeightValue));
                UpdateThumbPos(thumb, data);
            };

            thumb.PreviewMouseLeftButtonDown += (s, e) => {
                vm.SelectedRamdata = data;
            };

            UpdateThumbPos(thumb, data);
            RamCanvas.Children.Add(thumb);
        }
        private void UpdateThumbPos(Thumb t, RamData d)
        {
            Canvas.SetLeft(t, d.Column * CharWidth);
            Canvas.SetTop(t, d.Row * LineHeightValue);
        }
        private ControlTemplate CreateRamTemplate()
        {
            // 1. テンプレートの対象となる型を指定
            ControlTemplate template = new ControlTemplate(typeof(Thumb));

            // 2. Border要素を作成
            FrameworkElementFactory borderFactory = new FrameworkElementFactory(typeof(Border));
            borderFactory.Name = "brd";
            borderFactory.SetValue(Border.BackgroundProperty, new SolidColorBrush(Color.FromArgb(0x80, 0xFF, 0xFF, 0xFF)));
            borderFactory.SetValue(Border.BorderThicknessProperty, new Thickness(0));

            // 3. DataTriggerを設定 (IsSelectedプロパティを監視)
            DataTrigger trigger = new DataTrigger
            {
                Binding = new Binding("IsSelected"),
                Value = true
            };
            // 選択されたら赤色にする
            trigger.Setters.Add(new Setter(Border.BackgroundProperty, new SolidColorBrush(Color.FromArgb(0x66, 0xFF, 0x00, 0x00)), "brd"));

            // 4. テンプレートに組み立てる
            template.VisualTree = borderFactory;
            template.Triggers.Add(trigger);

            return template;
        }

        // テキスト領域の空白をクリックしたら選択解除
        private void MainEditor_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            // RAM枠（Thumb）以外の場所をクリックした場合
            if (e.OriginalSource is not Border && e.OriginalSource is not System.Windows.Shapes.Shape)
            {
                var vm = (MainViewModel)this.DataContext;
                vm.ClearSelection();
            }
        }

        // DataGridの右クリックメニューからの削除
        private void DeleteRam_Click(object sender, RoutedEventArgs e)
        {
            DeleteActiveRam();
        }


        private void DeleteActiveRam()
        {
            var vm = (MainViewModel)this.DataContext;
            var target = vm.SelectedRamdata;
            if (target == null) return;

            // Canvas上の対応するThumbを探して削除
            Thumb thumbToDelete = FindThumbFor(target);
            if (thumbToDelete != null)
            {
                RamCanvas.Children.Remove(thumbToDelete);
            }

            // リストから削除
            vm.RemoveSelectedRam();
        }
        private Thumb FindThumbFor(RamData data)
        {
            foreach (var child in RamCanvas.Children)
            {
                if (child is Thumb t && t.Tag == data)
                {
                    return t;
                }
            }
            return null;
        }

        private void AddRamFromCatalog_ContextMenu_Click(object sender, RoutedEventArgs e)
        {
            var vm = (MainViewModel)this.DataContext;
            if (vm.SelectedRamCatalog == null) return;

            // 現在のテキストエディタのカーソル位置（行・桁）を取得
            TextPointer tp = MainEditor.CaretPosition;

            // 行の計算
            int row = 0;
            Paragraph currentPara = tp.Paragraph;
            foreach (var block in MainEditor.Document.Blocks)
            {
                if (block == currentPara) break;
                row++;
            }

            // 桁（列）の計算：段落の先頭からのオフセット
            TextRange range = new TextRange(currentPara.ContentStart, tp);
            int col = range.Text.Length;

            // 配置実行
            var newRam = new RamData
            {
                Catalog = vm.SelectedRamCatalog,
                Format = vm.SelectedRamCatalog.Format,
                Row = row,
                Column = col
            };

            vm.RamdataList.Add(newRam);
            int length = GetFormatLength(vm, newRam.Format);
            CreateRamVisual(newRam, length);
        }
        private int GetFormatLength(WpfApp1.ViewModels.MainViewModel vm, string formatId)
        {
            // FormatList(マスター)の中から、Idが一致するものを探す
            var formatInfo = vm.FormatList.FirstOrDefault(f => f.Id == formatId);

            // 見つかればそのLength、見つからなければデフォルトで8桁とする
            return formatInfo?.Length ?? 8;
        }

        protected override void OnPreviewKeyDown(KeyEventArgs e)
        {
            base.OnPreviewKeyDown(e);

            // Deleteキーが押された時の処理
            if (e.Key == Key.Delete)
            {
                // 何かRAMが選択されているか確認
                var vm = (WpfApp1.ViewModels.MainViewModel)this.DataContext;
                if (vm.SelectedRamdata != null)
                {
                    // 削除を実行
                    DeleteActiveRam();
                    // 他のコントロール（RichTextBoxの文字消去など）にイベントを流さない
                    e.Handled = true;
                }
            }
        }
        
        private void MainEditor_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (RamCanvas == null) return;

            // スクロールした分だけ、Canvasを逆方向にずらす
            // これにより、Canvas上の四角がテキストと一緒に動いているように見えます
            RamCanvas.RenderTransform = new TranslateTransform(
                -e.HorizontalOffset,
                -e.VerticalOffset);
        }
    }
}
