using AvalonDock.Layout;
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

        //private void NewFile_Click(object sender, RoutedEventArgs e)
        //{
        //    MainEditor.Document = new FlowDocument();
        //    _currentFilePath = null;
        //}

        // JSON読み込みメニュー用
        private void OpenFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
            };

            if (dialog.ShowDialog() == true)
            {
                // ... ダイアログでパス取得 ...
                var service = new JsonEditorService();
                var projectData = service.LoadProjectFromJson(dialog.FileName);

                var vm = (MainViewModel)this.DataContext;
                vm.EditorTabs.Clear(); // 一旦全タブ削除

                foreach (var tabData in projectData.Tabs)
                {
                    // 1. ViewModelを作成
                    var tabVM = new DisplayEditorViewModel(vm) { DisplayNumber = tabData.Title };

                    // 2. RAMデータを復元してVMにセット
                    foreach (var ram in tabData.Rams)
                    {
                        tabVM.PlacedRams.Add(new RamItemViewModel(ram, vm));
                    }

                    // 3. タブを追加（これで EditorView が生成される）
                    vm.EditorTabs.Add(tabVM);

                    // 【注意】UI(RichTextBox)の復元は、EditorViewがロードされた後に行う必要があります
                    // ここは一工夫必要（後述）
                }
            }
        }
        private void SaveFile_Click(object sender, RoutedEventArgs e)
        {
            //    if (string.IsNullOrEmpty(_currentFilePath))
            //    {
            //        SaveAsFile_Click(sender, e);
            //        return;
            //    }

            //    var service = new JsonEditorService();
            //    var vm = (MainViewModel)DataContext;
            //    var data = service.CreateSaveData(MainEditor, vm.RamdataList);

            //    service.SaveToJson(data, _currentFilePath);
        }
        private void SaveAsFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
            };

            if (dialog.ShowDialog() == true)
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
                        var editorView = GetVisualChild<EditorView>(cp);
                        if (editorView != null)
                        {
                            saveData.Tabs.Add(editorView.GetEditorData());
                        }
                    }
                }
                service.SaveToJson(saveData, dialog.FileName);

                _currentFilePath = dialog.FileName;
            }
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
                    var editorView = GetVisualChild<EditorView>(cp);

                    if (editorView != null)
                    {
                        // 5. EditorView の色変更メソッドを呼ぶ
                        editorView.ApplyColorToSelection(index);
                    }
                }
            }
        }

        private T? GetVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            if (parent == null) return null;
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
