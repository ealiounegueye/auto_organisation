using System.Collections.ObjectModel;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Identity.Client;
using OutlookOrganizer.Configuration;
using OutlookOrganizer.Models;
using OutlookOrganizer.Services;

namespace OutlookOrganizer.ViewModels;

public sealed class MainViewModel : ViewModelBase
{
    private readonly AppOptions _options;
    private GraphAuthService? _auth;
    private GraphMailService? _mail;
    private bool _isBusy;
    private bool _isDone;
    private bool _isSignedIn;
    private bool _hasAnalysis;
    private string _headline = "Organisation de votre boîte";
    private string _statusMessage = "Connexion à Microsoft 365…";
    private string _account = "Non connecté";
    private int _analyzedCount;
    private int _classifiableCount;
    private int _unclassifiedCount;

    public MainViewModel(AppOptions options)
    {
        _options = options;
        CloseCommand = new RelayCommand(() => CloseRequested?.Invoke());
    }

    public event Action? CloseRequested;

    public ObservableCollection<CategoryStat> Categories { get; } = [];

    public ICommand CloseCommand { get; }

    public bool IsBusy
    {
        get => _isBusy;
        private set
        {
            if (SetProperty(ref _isBusy, value))
            {
                OnPropertyChanged(nameof(CanClose));
            }
        }
    }

    public bool IsDone
    {
        get => _isDone;
        private set
        {
            if (SetProperty(ref _isDone, value))
            {
                OnPropertyChanged(nameof(CanClose));
            }
        }
    }

    public bool CanClose => !IsBusy;

    public bool IsSignedIn
    {
        get => _isSignedIn;
        private set => SetProperty(ref _isSignedIn, value);
    }

    public bool HasAnalysis
    {
        get => _hasAnalysis;
        private set
        {
            if (SetProperty(ref _hasAnalysis, value))
            {
                OnPropertyChanged(nameof(SummaryText));
            }
        }
    }

    public string Headline
    {
        get => _headline;
        private set => SetProperty(ref _headline, value);
    }

    public string StatusMessage
    {
        get => _statusMessage;
        private set => SetProperty(ref _statusMessage, value);
    }

    public string Account
    {
        get => _account;
        private set => SetProperty(ref _account, value);
    }

    public int AnalyzedCount
    {
        get => _analyzedCount;
        private set
        {
            if (SetProperty(ref _analyzedCount, value))
            {
                OnPropertyChanged(nameof(SummaryText));
            }
        }
    }

    public int ClassifiableCount
    {
        get => _classifiableCount;
        private set
        {
            if (SetProperty(ref _classifiableCount, value))
            {
                OnPropertyChanged(nameof(SummaryText));
            }
        }
    }

    public int UnclassifiedCount
    {
        get => _unclassifiedCount;
        private set
        {
            if (SetProperty(ref _unclassifiedCount, value))
            {
                OnPropertyChanged(nameof(SummaryText));
            }
        }
    }

    public string SummaryText => HasAnalysis
        ? $"{AnalyzedCount:N0} messages lus — {Categories.Count(item => item.Name != "Non classés")} thématique(s) — {ClassifiableCount:N0} classés"
        : "L’application crée les thématiques à partir de vos messages.";

    public async Task RunAutomaticallyAsync()
    {
        if (!_options.HasValidClientId)
        {
            Headline = "Configuration manquante";
            StatusMessage = "L’administrateur doit enregistrer l’application dans Entra ID avant de distribuer ce fichier.";
            IsDone = true;
            return;
        }

        IsBusy = true;
        IsDone = false;
        HasAnalysis = false;
        Categories.Clear();
        Headline = "Organisation de votre boîte";
        StatusMessage = "Connexion à Microsoft 365…";

        try
        {
            EnsureServices();
            await _auth!.InitializeCacheAsync();

            await _auth.SignInAsync();
            await LoadProfileAsync();
            StatusMessage = "Lecture de votre boîte…";

            var progress = new Progress<string>(message => StatusMessage = message);
            var analysis = await _mail!.AnalyzeInboxAsync(progress);
            ShowAnalysis(analysis);

            if (_options.Organizer.SimulationModeDefault)
            {
                Headline = "Aperçu terminé";
                StatusMessage = analysis.Themes.Count == 0
                    ? "Pas assez de messages semblables pour créer une thématique. Rien n’a été modifié."
                    : $"{analysis.Themes.Count} thématique(s) seraient créées. Rien n’a été modifié (mode aperçu).";
                return;
            }

            if (analysis.Themes.Count == 0)
            {
                Headline = "Terminé";
                StatusMessage = "Pas assez de messages semblables pour créer une thématique. Rien n’a été modifié.";
                return;
            }

            StatusMessage = "Création des thématiques dans Outlook…";
            var applied = await _mail.ApplyThemesAsync(analysis, progress);
            Headline = "C’est terminé";
            StatusMessage = $"{analysis.Themes.Count} thématique(s) créées — {applied:N0} messages classés. Vous pouvez ouvrir Outlook, puis fermer cette fenêtre.";
        }
        catch (MsalClientException ex) when (ex.ErrorCode == "authentication_canceled")
        {
            Headline = "Connexion annulée";
            StatusMessage = "Relancez l’application et connectez-vous avec votre compte Microsoft 365.";
        }
        catch (Exception ex)
        {
            Headline = "Une erreur s’est produite";
            StatusMessage = HumanizeError(ex);
        }
        finally
        {
            IsBusy = false;
            IsDone = true;
        }
    }

    private void ShowAnalysis(MailboxAnalysis analysis)
    {
        AnalyzedCount = analysis.AnalyzedCount;
        ClassifiableCount = analysis.ClassifiableCount;
        UnclassifiedCount = analysis.UnclassifiedCount;
        Categories.Clear();

        foreach (var theme in analysis.Themes)
        {
            var brush = (SolidColorBrush)new BrushConverter().ConvertFromString(theme.ColorHex)!;
            if (brush.CanFreeze)
            {
                brush.Freeze();
            }

            Categories.Add(new CategoryStat
            {
                Name = theme.Name,
                Color = brush,
                Count = theme.Count,
                Source = theme.Source
            });
        }

        if (analysis.UnclassifiedCount > 0)
        {
            var other = new SolidColorBrush(Color.FromRgb(0x60, 0x5E, 0x5C));
            other.Freeze();
            Categories.Add(new CategoryStat
            {
                Name = "Non classés",
                Color = other,
                Count = analysis.UnclassifiedCount,
                Source = "pas assez de points communs"
            });
        }

        HasAnalysis = true;
    }

    private async Task LoadProfileAsync()
    {
        var profile = await _mail!.GetProfileAsync();
        Account = profile.Account;
        IsSignedIn = true;
    }

    private void EnsureServices()
    {
        _auth ??= new GraphAuthService(_options, GraphAuthService.GetActiveWindowHandle);
        _mail ??= new GraphMailService(_auth, _options);
    }

    private static string HumanizeError(Exception ex) => ex switch
    {
        InvalidOperationException => ex.Message,
        MsalServiceException msal when msal.Message.Contains("AADSTS65001", StringComparison.OrdinalIgnoreCase) =>
            "L'administrateur Microsoft 365 doit consentir aux permissions Graph.",
        MsalServiceException msal when msal.Message.Contains("AADSTS700016", StringComparison.OrdinalIgnoreCase) =>
            "Le Client ID Entra ID est incorrect, ou l'application n'existe pas dans votre locataire.",
        HttpRequestException => "Impossible de joindre Microsoft 365. Vérifiez le réseau et relancez.",
        _ => $"Erreur : {ex.Message}"
    };
}
