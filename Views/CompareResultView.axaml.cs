using System;
using System.Diagnostics;
using System.IO;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Avalonia.Media.Imaging;
using PixelCompareSuite.ViewModels;

namespace PixelCompareSuite.Views
{
    public partial class CompareResultView : Window
    {
        public CompareResultView()
        {
            InitializeComponent();
            this.Opened += CompareResultView_Opened;
        }

        private void InitializeComponent()
        {
            AvaloniaXamlLoader.Load(this);
        }

        private void CompareResultView_Opened(object? sender, EventArgs e)
        {
            if (DataContext is CompareResultViewModel viewModel)
            {
                var topLevel = TopLevel.GetTopLevel(this);
                if (topLevel != null)
                {
                    viewModel.SetTopLevel(topLevel);
                }
            }
        }

        private void OnItemPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            if (sender is Border border && border.DataContext is CompareItemViewModel item)
            {
                if (DataContext is CompareResultViewModel viewModel)
                {
                    viewModel.SelectItemCommand.Execute(item);
                }
            }
        }
        
        
        private void OnDiffImagePointerPressed(object? sender,PointerPressedEventArgs e)
        {
            if (e.ClickCount != 2)
                return;

            if (DataContext is not CompareResultViewModel vm)
                return;

            var path = vm.SelectedItem?.Image2BitmapPath;
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                return;

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = path,
                    UseShellExecute = true
                });
            }
            catch (Exception exception)
            {
                Debug.WriteLine(exception);
            }
        }
    }
}

