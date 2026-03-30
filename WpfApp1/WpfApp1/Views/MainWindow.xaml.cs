using AvalonDock.Layout;
using ExcelDataReader.Log;
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
using WpfApp1;
using WpfApp1.Helpers;
using WpfApp1.Models;
using WpfApp1.Services;
using WpfApp1.ViewModels;
using WpfApp1.Views;

namespace WpfApp1.Views
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        // 0～9用のコマンドを定義
        public static RoutedCommand ColorCommand = new RoutedCommand();

        public static readonly RoutedUICommand NewProjectCommand = new RoutedUICommand("新規作成", "NewProject", typeof(MainWindow),new InputGestureCollection { new KeyGesture(Key.N, ModifierKeys.Control) });
        public static readonly RoutedUICommand OpenCommand = new RoutedUICommand("開く", "Open", typeof(MainWindow), new InputGestureCollection { new KeyGesture(Key.O, ModifierKeys.Control) });
        public static readonly RoutedUICommand SaveCommand = new RoutedUICommand("上書き保存", "Save", typeof(MainWindow), new InputGestureCollection { new KeyGesture(Key.S, ModifierKeys.Control) });
        public static readonly RoutedUICommand SaveAsCommand = new RoutedUICommand("名前を付けて保存", "SaveAs", typeof(MainWindow), new InputGestureCollection { new KeyGesture(Key.S, ModifierKeys.Control | ModifierKeys.Shift) });
        public static readonly RoutedUICommand ExitCommand = new RoutedUICommand("終了", "Exit", typeof(MainWindow),new InputGestureCollection { new KeyGesture(Key.F4, ModifierKeys.Alt) });

        public static readonly RoutedUICommand NewTabCommand = new RoutedUICommand("新しいタブ", "NewTab", typeof(MainWindow));
        public static readonly RoutedUICommand AddRamCommand = new RoutedUICommand("RAM配置", "AddRam", typeof(MainWindow), new InputGestureCollection { new KeyGesture(Key.Insert, ModifierKeys.Control | ModifierKeys.Shift) });
        public static readonly RoutedUICommand DeleteRamCommand = new RoutedUICommand("RAM削除", "DeleteRam", typeof(MainWindow), new InputGestureCollection { new KeyGesture(Key.Delete) });



        public MainWindow()
        {
            InitializeComponent();
            // ViewModelをセット
            var vm = new MainViewModel();
            this.DataContext = vm;
            // コマンドと既存メソッドの紐付け（CommandBinding）
            this.CommandBindings.Add(new CommandBinding(NewProjectCommand, NewProject_Click));
            this.CommandBindings.Add(new CommandBinding(ColorCommand, ColorCommand_Executed));
            this.CommandBindings.Add(new CommandBinding(OpenCommand, OpenFile_Click));
            this.CommandBindings.Add(new CommandBinding(SaveCommand, SaveFile_Click));
            this.CommandBindings.Add(new CommandBinding(SaveAsCommand, SaveAsFile_Click));
            this.CommandBindings.Add(new CommandBinding(ExitCommand, (s, e) => this.Close()));

            this.CommandBindings.Add(new CommandBinding(NewTabCommand, (s, e) => {
                vm.AddTabCommand.Execute(null);
            }));
            this.CommandBindings.Add(new CommandBinding(AddRamCommand, (s, e) => {
                if (vm.AddRamFromCatalogCommand.CanExecute(null))
                    vm.AddRamFromCatalogCommand.Execute(null);
            }));
            this.CommandBindings.Add(new CommandBinding(DeleteRamCommand, (s, e) => {
                if (vm.RemoveRamCommand.CanExecute(null))
                    vm.RemoveRamCommand.Execute(null);
            }));
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            base.OnClosing(e);

            // 以前作成した「変更があれば保存確認する」メソッドを呼び出す
            if (!ConfirmSaveIfDirty())
            {
                // キャンセルされたら、終了自体を取り消す
                e.Cancel = true;
            }
        }


        private void ShowFormatPane_Click(object sender, RoutedEventArgs e)
        {
            // AvalonDockの機能で、非表示状態から再表示させます
            if (FormatPane != null)
            {
                FormatPane.IsVisible = true;
                FormatPane.IsSelected = true;
            }
        }

        private void ShowRamCatalogPane_Click(object sender, RoutedEventArgs e)
        {
            if (RamCatalogPane != null)
            {
                RamCatalogPane.IsVisible = true;
                RamCatalogPane.IsSelected = true;
            }
        }
        
        private void ShowRamDatePanePane_Click(object sender, RoutedEventArgs e)
        {
            if (RamDatePane != null)
            {
                RamDatePane.IsVisible = true;
                RamDatePane.IsSelected = true;
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

        private void NewProject_Click(object sender, ExecutedRoutedEventArgs e)
        {
            if (!ConfirmSaveIfDirty()) return; // 変更あれば保存確認（以前作成したメソッド）

            var vm = (MainViewModel)this.DataContext;

            // 1. 全データをクリア
            vm.EditorTabs.Clear();
            vm.CurrentFilePath = null;

            // 2. 初期タブを1枚だけ作成
            vm.AddTabCommand.Execute(null);

            // 3. 状態リセット
            vm.IsDirty = false;
        }

        private bool ConfirmSaveIfDirty()
        {
            var vm = (MainViewModel)DataContext;
            if (!vm.IsDirty) return true; // 変更なければ次へ

            var result = MessageBox.Show("変更を保存しますか？", "確認",
                MessageBoxButton.YesNoCancel, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                SaveFile_Click(null, null); // 保存実行
                return !vm.IsDirty; // 保存成功ならTrue
            }

            return result == MessageBoxResult.No; // 「いいえ」なら破棄OK
        }

        // JSON読み込みメニュー用
        private async void OpenFile_Click(object sender, RoutedEventArgs e)
        {
            if (!ConfirmSaveIfDirty()) return; // キャンセルなら中断

            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
            };

            // ダイアログでパス取得
            if (dialog.ShowDialog() == true)
            {
                var service = new JsonEditorService();
                var projectData = service.LoadProjectFromJson(dialog.FileName);

                var mainVm = (MainViewModel)this.DataContext;
                mainVm.CurrentFilePath = dialog.FileName;
                mainVm.EditorTabs.Clear(); // 一旦全タブ削除

                foreach (var tabData in projectData.Tabs)
                {
                    // ViewModelを作成
                    var tabVM = new DisplayEditorViewModel(mainVm) { DisplayName = tabData.Title };

                    // RAMデータを復元してVMにセット
                    foreach (var ram in tabData.Rams)
                    {
                        tabVM.PlacedRams.Add(new RamItemViewModel(ram, mainVm));
                    }
                    // EditorViewのLoadedイベントで、レイアウトやシンボルの復元を行うためのデータを渡す
                    tabVM.RestoreData = tabData;
                    // タブを追加（これで EditorView が生成される）
                    mainVm.EditorTabs.Add(tabVM);
                }

                // 全てのタブを追加し終わった後
                foreach (var tab in mainVm.EditorTabs)
                {
                    // 1. プログラムからタブを切り替える
                    mainVm.ActiveTab = tab;

                    // 2. 一瞬待機して、WPFが EditorView を生成・Loadedイベントを走らせる時間を稼ぐ
                    // DispatcherPriority.Background を使うことで描画更新を待てます
                    await Dispatcher.BeginInvoke(new Action(() => { }), System.Windows.Threading.DispatcherPriority.Background);
                }

                mainVm.IsDirty = false;
            }
        }
        private void SaveFile_Click(object sender, RoutedEventArgs e)
        {
            var mainVm = (MainViewModel)this.DataContext;
            if (string.IsNullOrEmpty(mainVm.CurrentFilePath))
            {
                SaveAsFile_Click(sender, e);
                return;
            }

            SaveFile(mainVm.CurrentFilePath);
        }
        private void SaveAsFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
            };

            if (dialog.ShowDialog() == true)
            {
                var mainVm = (MainViewModel)this.DataContext;
                mainVm.CurrentFilePath = dialog.FileName;
                SaveFile(mainVm.CurrentFilePath);
            }
        }

        private void SaveFile(string path)
        {
            var saveData = new ProjectSaveData();
            var service = new JsonEditorService();

            var documents = dockingManager.Layout.Descendents().OfType<AvalonDock.Layout.LayoutDocument>();
            foreach (var doc in documents)
            {
                var layoutItem = dockingManager.GetLayoutItemFromModel(doc);
                if (layoutItem?.View is ContentPresenter cp)
                {
                    // 各タブの実体(EditorView)からデータを取得
                    var editorView = VisualTreeHelperExtensions.GetVisualChild<EditorView>(cp);
                    if (editorView != null)
                    {
                        saveData.Tabs.Add(editorView.GetEditorData());
                    }
                }
            }
            service.SaveToJson(saveData, path);
            var mainVm = (MainViewModel)this.DataContext;
            mainVm.IsDirty = false;
        }

        private void ColorCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            if (e.Parameter is string index)
            {
                // 1. アクティブなコンテンツ（ViewModel）を取得
                var activeContent = dockingManager.ActiveContent;
                if (activeContent == null) return;

                // 2. 全ドキュメントの中から、このViewModelを持っている「実体」を探す
                // LayoutDocumentPane内の全ドキュメントをスキャンします
                var activeLayoutDocument = dockingManager.Layout.Descendents()
                    .OfType<AvalonDock.Layout.LayoutDocument>()
                    .FirstOrDefault(d => d.Content == activeContent);

                if (activeLayoutDocument == null) return;

                // 3. レイアウトアイテム（枠組み）を取得
                var layoutItem = dockingManager.GetLayoutItemFromModel(activeLayoutDocument);
                if (layoutItem?.View is ContentPresenter cp)
                {
                    // 4. 【ここが重要】ContentPresenter の「視覚的な子要素」から EditorView を探す
                    // cp.Content は ViewModel を指しているため、実体（EditorView）を VisualTree から掘り起こす
                    var editorView = VisualTreeHelperExtensions.GetVisualChild<EditorView>(cp);

                    if (editorView != null)
                    {
                        // 5. EditorView の色変更メソッドを呼ぶ
                        editorView.ApplyColorToSelection(index);
                    }
                }
            }
        }

        private void dockingManager_DocumentClosing(object sender, AvalonDock.DocumentClosingEventArgs e)
        {
            // 閉じようとしているタブの ViewModel を取得
            var tabVM = e.Document.Content as DisplayEditorViewModel;
            if (tabVM == null) return;

            // 確認メッセージを表示
            var result = MessageBox.Show(
                $"{tabVM.DisplayName} を閉じますか？\n（タブの内容は破棄されます。）",
                "タブを閉じる確認",
                MessageBoxButton.OKCancel,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Cancel)
            {
                // 1. キャンセルなら、閉じる動作自体を無効にする
                e.Cancel = true;
            }
            else
            {
                // 2. OKなら、ViewModel のリスト(EditorTabs)からも削除する
                // これをしないと、画面からは消えてもメモリ（保存対象）に残ってしまいます
                var mainVM = (MainViewModel)this.DataContext;
                mainVM.EditorTabs.Remove(tabVM);

                if (mainVM.EditorTabs.Count == 0) mainVM.IsDirty = false;
            }
        }

        //private T? GetVisualChild<T>(DependencyObject parent) where T : DependencyObject
        //{
        //    if (parent == null) return null;
        //    for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
        //    {
        //        var child = VisualTreeHelper.GetChild(parent, i);
        //        if (child is T t) return t;
        //        var result = GetVisualChild<T>(child);
        //        if (result != null) return result;
        //    }
        //    return null;
        //}





    }

}
