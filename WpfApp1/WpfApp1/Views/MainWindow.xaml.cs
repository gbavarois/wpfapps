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
                var vm = (MainViewModel)this.DataContext;

                // Command.Execute(...) ではなく、直接メソッドを呼ぶ
                var saveData = vm.LoadFromJson(dialog.FileName);

                if (saveData != null)
                {
                    var service = new JsonEditorService();
                    //service.RestoreText(MainEditor, saveData.Lines);        // エラーを消すため暫定的にコメント化する
                    //service.ApplyColorInfo(MainEditor, saveData.Colors);    // エラーを消すため暫定的にコメント化する
                }

                _currentFilePath = dialog.FileName;
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
            //    var dialog = new SaveFileDialog
            //    {
            //        Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
            //    };

            //    if (dialog.ShowDialog() == true)
            //    {
            //        var service = new JsonEditorService();
            //        var vm = (MainViewModel)DataContext;
            //        var data = service.CreateSaveData(MainEditor, vm.RamdataList);

            //        service.SaveToJson(data, dialog.FileName);

            //        _currentFilePath = dialog.FileName;
            //    }
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

                // 4. ContentPresenterの中身から EditorView を取り出す
                if (layoutItem?.View is ContentPresenter cp && cp.Content is EditorView editorView)
                {
                    // 5. EditorView の色変更メソッドを呼ぶ
                    editorView.ApplyColorToSelection(index);
                }
            }
        }

        

        

    }

}
