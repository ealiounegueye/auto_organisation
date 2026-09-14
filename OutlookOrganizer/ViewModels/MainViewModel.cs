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
    private string _statusMessage = "Connectez-vous avec votre compte Microsoft 365 pour commencer.";
    private string _account = "Non connecté";
    private string _displayName = string.Empty;
    private int _analyzedCount;
    private int _classifiableCount;

    public MainViewModel(AppOptions options)
    {
        _options = options;
        _isSimulationMode = options.Organizer.SimulationModeDefault;

        foreach (var category in MailCategory.All)
        {
            var brush = (SolidColorBrush)new BrushConverter().ConvertFromString(category.ColorKey)!;
            if (brush.CanFreeze)
            {
                brush.Freeze();
            }

            Categories.Add(new CategoryStat
            {
                Name = category.Name,
                Color = brush
            });
        }

        SignInCommand = new AsyncRelayCommand(SignInAsync, () => !IsBusy);
        SignOutCommand = new AsyncRelayCommand(SignOutAsync, () => !IsBusy && IsSignedIn);
        AnalyzeCommand = new AsyncRelayCommand(AnalyzeAsync, () => !IsBusy && IsSignedIn);
        OrganizeCommand = new AsyncRelayCommand(OrganizeAsync, () => !IsBusy && IsSignedIn && HasAnalysis);
    }

    public ObservableCollection<CategoryStat> Categories { get; } = [];

    public ICommand SignInCommand { get; }
    public ICommand SignOutCommand { get; }
    public ICommand AnalyzeCommand { get; }
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
                RefreshCommands();
            }
        }
    }

    public bool IsSimulationMode
    {
        get => _isSimulationMode;
        set
        {
            if (SetProperty(ref _isSimulationMode, value))
            {
                OnPropertyChanged(nameof(SimulationBannerText));
            }
        }
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

    public string SummaryText => HasAnalysis
        ? $"{AnalyzedCount:N0} messages analysés — {ClassifiableCount:N0} peuvent être classés"
        : "Aucune analyse pour le moment";

    public string SimulationBannerText => IsSimulationMode
        ? "Mode simulation activé — aucun message ne sera modifié."
        : "Mode réel — les catégories Outlook seront écrites (disponible à l'étape 2).";

    public async Task InitializeAsync()
    {
        if (!_options.HasValidClientId)
        {
            StatusMessage = "Configurez d'abord le Client ID dans appsettings.local.json (voir le guide Entra ID).";
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
            StatusMessage = "Session Microsoft 365 restaurée. Vous pouvez analyser votre boîte.";
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
            StatusMessage = "Le Client ID Entra ID n'est pas renseigné. Ouvrez docs/ENTRA-ID.md puis appsettings.local.json.";
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
            StatusMessage = "Connexion réussie. Cliquez sur Analyser ma boîte — rien ne sera modifié.";
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
            ResetCategoryCounts();
            StatusMessage = "Déconnecté. Vos identifiants n'ont jamais été stockés par l'application.";
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

    private async Task AnalyzeAsync()
    {
        if (_mail is null)
        {
            StatusMessage = "Connectez-vous avant d'analyser la boîte.";
            return;
        }

        IsBusy = true;
        HasAnalysis = false;
        StatusMessage = "Lecture de la boîte (simulation, lecture seule)…";

        try
        {
            var progress = new Progress<string>(message => StatusMessage = message);
            var analysis = await _mail.AnalyzeInboxAsync(progress);

            AnalyzedCount = analysis.AnalyzedCount;
            ClassifiableCount = analysis.ClassifiableCount;
            ApplyCounts(analysis.CategoryCounts);
            HasAnalysis = true;
            StatusMessage = $"{analysis.AnalyzedCount:N0} messages analysés — {analysis.ClassifiableCount:N0} peuvent être classés. Aucun message n'a été modifié.";
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

    private Task OrganizeAsync()
    {
        if (IsSimulationMode)
        {
            StatusMessage = $"Simulation : {ClassifiableCount:N0} messages seraient classés. Aucun message n'a été modifié.";
            return Task.CompletedTask;
        }

        StatusMessage = "Le classement réel (écriture des catégories Outlook) arrive à l'étape 2. Laissez le mode simulation activé pour l'instant.";
        return Task.CompletedTask;
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

    private void ApplyCounts(IReadOnlyDictionary<string, int> counts)
    {
        foreach (var category in Categories)
        {
            category.Count = counts.TryGetValue(category.Name, out var value) ? value : 0;
        }
    }

    private void ResetCategoryCounts()
    {
        foreach (var category in Categories)
        {
            category.Count = 0;
        }
    }

    private void RefreshCommands()
    {
        (SignInCommand as AsyncRelayCommand)?.RaiseCanExecuteChanged();
        (SignOutCommand as AsyncRelayCommand)?.RaiseCanExecuteChanged();
        (AnalyzeCommand as AsyncRelayCommand)?.RaiseCanExecuteChanged();
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
