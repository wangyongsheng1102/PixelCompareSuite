using System;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
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
                viewModel.SetTopLevel(TopLevel.GetTopLevel(this));
            }
        }

        private void OnItemPointerPressed(object? sender, RoutedEventArgs e)
        {
            if (sender is Border border && border.DataContext is CompareItemViewModel item)
            {
                if (DataContext is CompareResultViewModel viewModel)
                {
                    viewModel.SelectItemCommand.Execute(item);
                }
            }
        }
    }
}

