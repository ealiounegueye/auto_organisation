using System.Windows;
using Microsoft.Extensions.Configuration;
using OutlookOrganizer.Configuration;
using OutlookOrganizer.ViewModels;

namespace OutlookOrganizer;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        var configuration = new ConfigurationBuilder()
            .SetBasePath(AppContext.BaseDirectory)
            .AddJsonFile("appsettings.json", optional: false)
            .AddJsonFile("appsettings.local.json", optional: true)
            .Build();

        var options = new AppOptions();
        configuration.Bind(options);

        var viewModel = new MainViewModel(options);
        var window = new MainWindow(viewModel);
        window.Show();
    }
}
