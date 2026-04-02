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
        private MainViewModel vm => (MainViewModel)DataContext;


        public MainWindow()
        {
            InitializeComponent();
            //// ViewModelをセット
            var vm = new MainViewModel();
            this.DataContext = vm;

            Loaded += (s, e) =>
            {
                var vm = (MainViewModel)DataContext;

                vm.RequestSaveBeforeContinue += ConfirmSaveIfDirtyAsync;
                vm.RequestOpenFilePath += ShowOpenDialogAsync;
                vm.RequestSaveFilePath += ShowSaveDialogAsync;
                vm.SaveRequested += SaveFile;
            };

            this.Closing += MainWindow_Closing;
        }

        //protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        //{
        //    base.OnClosing(e);

        //    // 以前作成した「変更があれば保存確認する」メソッドを呼び出す
        //    if (!ConfirmSaveIfDirty())
        //    {
        //        // キャンセルされたら、終了自体を取り消す
        //        e.Cancel = true;
        //    }
        //}


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
                //var vm = (MainViewModel)this.DataContext;
                // ViewModelのコマンドを、パスを引数にして実行する
                vm.LoadExcelCommand.Execute(dialog.FileName);
            }
        }

        private async Task<bool> ConfirmSaveIfDirtyAsync()
        {
            var vm = (MainViewModel)DataContext;

            if (!vm.IsDirty) return true;

            var result = MessageBox.Show(
                "変更を保存しますか？",
                "確認",
                MessageBoxButton.YesNoCancel);

            if (result == MessageBoxResult.Yes)
            {
                return await SaveFileInternal(); // ← Command使わない！
            }

            return result == MessageBoxResult.No;
        }

        //private bool SaveWithDialog()
        //{
        //    if (!string.IsNullOrEmpty(vm.CurrentFilePath))
        //    {
        //        vm.SaveFileCommand.Execute(vm.CurrentFilePath);
        //        return true;
        //    }

        //    return ShowSaveAsDialog();
        //}



        private Task<string?> ShowOpenDialogAsync()
        {
            var dialog = new OpenFileDialog
            {
                Filter = "JSON files (*.json)|*.json"
            };

            return Task.FromResult(dialog.ShowDialog() == true ? dialog.FileName : null);
        }

        private Task<string?> ShowSaveDialogAsync()
        {
            var dialog = new SaveFileDialog
            {
                Filter = "JSON files (*.json)|*.json"
            };

            return Task.FromResult(dialog.ShowDialog() == true ? dialog.FileName : null);
        }

        private async Task<bool> SaveFileInternal()
        {
            var vm = (MainViewModel)DataContext;

            if (string.IsNullOrEmpty(vm.CurrentFilePath))
            {
                var path = await ShowSaveDialogAsync();
                if (string.IsNullOrEmpty(path)) return false;

                vm.CurrentFilePath = path;
            }

            SaveFile(vm.CurrentFilePath);
            return true;
        }

        private void SaveFile(string path)
        {
            var saveData = new ProjectSaveData();
            var service = new JsonEditorService();

            var documents = dockingManager.Layout.Descendents()
                .OfType<AvalonDock.Layout.LayoutDocument>();

            foreach (var doc in documents)
            {
                var layoutItem = dockingManager.GetLayoutItemFromModel(doc);
                if (layoutItem?.View is ContentPresenter cp)
                {
                    var editorView = VisualTreeHelperExtensions.GetVisualChild<EditorView>(cp);
                    if (editorView != null)
                    {
                        saveData.Tabs.Add(editorView.GetEditorData());
                    }
                }
            }

            service.SaveToJson(saveData, path);

            var vm = (MainViewModel)DataContext;
            vm.IsDirty = false;
        }

        private async void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            var vm = (MainViewModel)DataContext;

            if (!vm.IsDirty) return;

            var result = MessageBox.Show(
                "変更を保存しますか？",
                "確認",
                MessageBoxButton.YesNoCancel);

            if (result == MessageBoxResult.Cancel)
            {
                e.Cancel = true; // ← 閉じるの中断
                return;
            }

            if (result == MessageBoxResult.Yes)
            {
                var success = await SaveAsInternal();

                if (!success)
                {
                    e.Cancel = true; // ← 保存キャンセルされたら閉じない
                }
            }
        }

        private async Task<bool> SaveAsInternal()
        {
            var vm = (MainViewModel)DataContext;

            // ★必ずダイアログ出す（既存パス無視）
            var path = await ShowSaveDialogAsync();
            if (string.IsNullOrEmpty(path)) return false;

            vm.CurrentFilePath = path;

            SaveFile(path);

            return true;
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

        private void RamCatalogList_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed && sender is DataGrid grid)
            {
                // 選択されているアイテム（RamCatalog）を取得
                var selectedItem = grid.SelectedItem as RamCatalog;
                if (selectedItem != null)
                {
                    var vm = (MainViewModel)this.DataContext;
                    vm.IsDraggingCatalog = true; // ★ ドラッグ開始（全Canvasを実体化）
                    //// ドラッグ操作を開始（データを DataObject に詰める）
                    //DataObject data = new DataObject("RamCatalogData", selectedItem);
                    //DragDrop.DoDragDrop(grid, data, DragDropEffects.Copy);
                    try
                    {
                        DataObject data = new DataObject("RamCatalogData", selectedItem);
                        // ★DoDragDropはドロップが完了するかキャンセルされるまでここでブロック（待機）します
                        DragDrop.DoDragDrop(grid, data, DragDropEffects.Copy);
                    }
                    finally
                    {
                        // 2. ドロップ完了、またはキャンセル！全タブのCanvasを透過状態(null)に戻す
                        vm.IsDraggingCatalog = false;
                    }
                }
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
