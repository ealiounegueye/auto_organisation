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

        var preview = e.Args.Any(argument =>
            argument.Equals("--preview", StringComparison.OrdinalIgnoreCase));

        var configuration = new ConfigurationBuilder()
            .SetBasePath(FindConfigDirectory())
            .AddJsonFile("appsettings.json", optional: false)
            .AddJsonFile("appsettings.local.json", optional: true)
            .Build();

        var options = new AppOptions();
        configuration.Bind(options);
        if (preview)
        {
            options.Organizer.SimulationModeDefault = true;
        }

        var viewModel = new MainViewModel(options);
        var window = new MainWindow(viewModel);
        window.Show();
    }

    private static string FindConfigDirectory()
    {
        var candidates = new[]
        {
            Path.GetDirectoryName(Environment.ProcessPath),
            AppContext.BaseDirectory,
            Directory.GetCurrentDirectory()
        };

        foreach (var directory in candidates)
        {
            if (!string.IsNullOrWhiteSpace(directory)
                && File.Exists(Path.Combine(directory, "appsettings.json")))
            {
                return directory;
            }
        }

        return AppContext.BaseDirectory;
    }
}
