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
    private bool _isSignedIn;
    private bool _hasAnalysis;
    private bool _isSimulationMode;
    private string _statusMessage = "Cliquez sur Organiser ma boîte. Les thématiques seront créées à partir de vos messages.";
    private string _account = "Non connecté";
    private string _displayName = string.Empty;
    private int _analyzedCount;
    private int _classifiableCount;
    private int _unclassifiedCount;

    public MainViewModel(AppOptions options)
    {
        _options = options;
        _isSimulationMode = options.Organizer.SimulationModeDefault;

        SignInCommand = new AsyncRelayCommand(SignInAsync, () => !IsBusy);
        SignOutCommand = new AsyncRelayCommand(SignOutAsync, () => !IsBusy && IsSignedIn);
        OrganizeCommand = new AsyncRelayCommand(OrganizeMailboxAsync, () => !IsBusy);
    }

    public ObservableCollection<CategoryStat> Categories { get; } = [];

    public ICommand SignInCommand { get; }
    public ICommand SignOutCommand { get; }
    public ICommand OrganizeCommand { get; }

    public bool IsBusy
    {
        get => _isBusy;
        private set
        {
            if (SetProperty(ref _isBusy, value))
            {
                RefreshCommands();
            }
        }
    }

    public bool IsSignedIn
    {
        get => _isSignedIn;
        private set
        {
            if (SetProperty(ref _isSignedIn, value))
            {
                OnPropertyChanged(nameof(IsSignedOut));
                RefreshCommands();
            }
        }
    }

    public bool IsSignedOut => !IsSignedIn;

    public bool HasAnalysis
    {
        get => _hasAnalysis;
        private set
        {
            if (SetProperty(ref _hasAnalysis, value))
            {
                OnPropertyChanged(nameof(HasNoThemesYet));
                OnPropertyChanged(nameof(SummaryText));
            }
        }
    }

    public bool HasNoThemesYet => !HasAnalysis;

    public bool IsSimulationMode
    {
        get => _isSimulationMode;
        set
        {
            if (SetProperty(ref _isSimulationMode, value))
            {
                OnPropertyChanged(nameof(SimulationBannerText));
                OnPropertyChanged(nameof(OrganizeButtonText));
            }
        }
    }

    public string OrganizeButtonText => IsSimulationMode
        ? "Aperçu de ma boîte"
        : "Organiser ma boîte";

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

    public string DisplayName
    {
        get => _displayName;
        private set => SetProperty(ref _displayName, value);
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
        ? $"{AnalyzedCount:N0} messages lus — {Categories.Count} thématique(s) créée(s) — {ClassifiableCount:N0} classés, {UnclassifiedCount:N0} laissés tels quels"
        : "Aucune thématique pour l’instant : elles naîtront de vos e-mails, pas d’une liste imposée.";

    public string SimulationBannerText => IsSimulationMode
        ? "Aperçu seulement — l’application lit votre boîte et propose des thématiques, sans rien modifier."
        : "Mode réel — un clic crée les thématiques dans Outlook et classe les messages correspondants.";

    public async Task InitializeAsync()
    {
        if (!_options.HasValidClientId)
        {
            StatusMessage = "L’administrateur doit d’abord enregistrer l’application dans Entra ID (docs/ENTRA-ID.md).";
            return;
        }

        try
        {
            EnsureServices();
            await _auth!.InitializeCacheAsync();
            var silent = await _auth.TrySignInSilentlyAsync();
            if (silent is null)
            {
                return;
            }

            await LoadProfileAsync();
            StatusMessage = "Session Microsoft 365 restaurée. Cliquez sur Organiser ma boîte.";
        }
        catch (Exception ex)
        {
            StatusMessage = HumanizeError(ex);
        }
    }

    private async Task SignInAsync()
    {
        if (!_options.HasValidClientId)
        {
            StatusMessage = "Le Client ID Entra ID n'est pas renseigné. Voir docs/ENTRA-ID.md.";
            return;
        }

        IsBusy = true;
        StatusMessage = "Ouverture de la connexion Microsoft 365…";

        try
        {
            EnsureServices();
            await _auth!.InitializeCacheAsync();
            await _auth.SignInAsync();
            await LoadProfileAsync();
            StatusMessage = "Connexion réussie. Cliquez sur Organiser ma boîte.";
        }
        catch (MsalClientException ex) when (ex.ErrorCode == "authentication_canceled")
        {
            StatusMessage = "Connexion annulée.";
        }
        catch (Exception ex)
        {
            StatusMessage = HumanizeError(ex);
        }
        finally
        {
            IsBusy = false;
        }
    }

    private async Task SignOutAsync()
    {
        IsBusy = true;
        try
        {
            if (_auth is not null)
            {
                await _auth.SignOutAsync();
            }

            IsSignedIn = false;
            HasAnalysis = false;
            Account = "Non connecté";
            DisplayName = string.Empty;
            AnalyzedCount = 0;
            ClassifiableCount = 0;
            UnclassifiedCount = 0;
            Categories.Clear();
            StatusMessage = "Déconnecté. Cliquez sur Organiser ma boîte pour recommencer.";
        }
        catch (Exception ex)
        {
            StatusMessage = HumanizeError(ex);
        }
        finally
        {
            IsBusy = false;
        }
    }

    private async Task OrganizeMailboxAsync()
    {
        if (!_options.HasValidClientId)
        {
            StatusMessage = "Le Client ID Entra ID n'est pas renseigné. Voir docs/ENTRA-ID.md.";
            return;
        }

        IsBusy = true;
        HasAnalysis = false;
        Categories.Clear();

        try
        {
            EnsureServices();
            await _auth!.InitializeCacheAsync();

            if (!IsSignedIn)
            {
                StatusMessage = "Ouverture de la connexion Microsoft 365…";
                await _auth.SignInAsync();
                await LoadProfileAsync();
            }

            StatusMessage = "Lecture de votre boîte…";
            var progress = new Progress<string>(message => StatusMessage = message);
            var analysis = await _mail!.AnalyzeInboxAsync(progress);
            ShowAnalysis(analysis);

            if (IsSimulationMode)
            {
                StatusMessage = analysis.Themes.Count == 0
                    ? "Aperçu terminé : pas assez de messages semblables pour créer une thématique."
                    : $"Aperçu : {analysis.Themes.Count} thématique(s) seraient créées, {analysis.ClassifiableCount:N0} messages classés. Rien n’a été modifié.";
                return;
            }

            if (analysis.Themes.Count == 0)
            {
                StatusMessage = "Pas assez de messages semblables pour créer une thématique. Rien n’a été modifié.";
                return;
            }

            StatusMessage = "Création des thématiques dans Outlook…";
            var applied = await _mail.ApplyThemesAsync(analysis, progress);
            StatusMessage = $"{analysis.Themes.Count} thématique(s) créées dans Outlook — {applied:N0} messages classés. Ouvrez Outlook pour les voir.";
        }
        catch (MsalClientException ex) when (ex.ErrorCode == "authentication_canceled")
        {
            StatusMessage = "Connexion annulée.";
        }
        catch (Exception ex)
        {
            StatusMessage = HumanizeError(ex);
        }
        finally
        {
            IsBusy = false;
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
        DisplayName = profile.DisplayName;
        Account = profile.Account;
        IsSignedIn = true;
    }

    private void EnsureServices()
    {
        _auth ??= new GraphAuthService(_options, GraphAuthService.GetActiveWindowHandle);
        _mail ??= new GraphMailService(_auth, _options);
    }

    private void RefreshCommands()
    {
        (SignInCommand as AsyncRelayCommand)?.RaiseCanExecuteChanged();
        (SignOutCommand as AsyncRelayCommand)?.RaiseCanExecuteChanged();
        (OrganizeCommand as AsyncRelayCommand)?.RaiseCanExecuteChanged();
    }

    private static string HumanizeError(Exception ex) => ex switch
    {
        InvalidOperationException => ex.Message,
        MsalServiceException msal when msal.Message.Contains("AADSTS65001", StringComparison.OrdinalIgnoreCase) =>
            "L'administrateur Microsoft 365 doit consentir aux permissions Graph (Mail.ReadWrite, MailboxSettings.ReadWrite).",
        MsalServiceException msal when msal.Message.Contains("AADSTS700016", StringComparison.OrdinalIgnoreCase) =>
            "Le Client ID Entra ID est incorrect, ou l'application n'existe pas dans votre locataire.",
        HttpRequestException => "Impossible de joindre Microsoft 365. Vérifiez le réseau et réessayez.",
        _ => $"Erreur : {ex.Message}"
    };
}
