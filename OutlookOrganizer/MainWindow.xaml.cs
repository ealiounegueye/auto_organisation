using System.Windows;
using OutlookOrganizer.ViewModels;

namespace OutlookOrganizer;

public partial class MainWindow : Window
{
    private readonly MainViewModel _viewModel;

    public MainWindow(MainViewModel viewModel)
    {
        _viewModel = viewModel;
        DataContext = _viewModel;
        _viewModel.CloseRequested += Close;
        InitializeComponent();
    }

    private async void OnLoaded(object sender, RoutedEventArgs e)
    {
        await _viewModel.RunAutomaticallyAsync();
    }
}
