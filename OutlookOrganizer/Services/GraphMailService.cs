using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions.Authentication;
using OutlookOrganizer.Configuration;
using OutlookOrganizer.Models;

namespace OutlookOrganizer.Services;

public sealed class GraphMailService
{
    private readonly GraphServiceClient _graph;
    private readonly AppOptions _options;

    public GraphMailService(GraphAuthService auth, AppOptions options)
    {
        _graph = new GraphServiceClient(new BaseBearerTokenAuthenticationProvider(new MsalTokenProvider(auth)));
        _options = options;
    }

    public async Task<GraphUserProfile> GetProfileAsync(CancellationToken cancellationToken = default)
    {
        var me = await _graph.Me.GetAsync(request =>
        {
            request.QueryParameters.Select = ["displayName", "mail", "userPrincipalName"];
        }, cancellationToken);

        var account = FirstNonEmpty(me?.Mail, me?.UserPrincipalName) ?? "Compte Microsoft 365";
        var displayName = FirstNonEmpty(me?.DisplayName, account) ?? account;

        return new GraphUserProfile
        {
            DisplayName = displayName,
            Account = account
        };
    }

    public async Task<MailboxAnalysis> AnalyzeInboxAsync(
        IProgress<string>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var profile = await GetProfileAsync(cancellationToken);
        var snapshots = new List<MailSnapshot>();
        var maxMessages = Math.Max(50, _options.Organizer.MaxMessagesToAnalyze);

        var page = await _graph.Me.Messages.GetAsync(request =>
        {
            request.QueryParameters.Select = ["id", "subject", "from", "categories", "hasAttachments", "receivedDateTime"];
            request.QueryParameters.Top = 100;
        }, cancellationToken);

        if (page is not null)
        {
            var iterator = PageIterator<Message, MessageCollectionResponse>.CreatePageIterator(
                _graph,
                page,
                message =>
                {
                    if (!string.IsNullOrWhiteSpace(message.Id))
                    {
                        snapshots.Add(ToSnapshot(message));
                    }

                    progress?.Report($"{snapshots.Count:N0} messages lus…");
                    return snapshots.Count < maxMessages;
                });

            await iterator.IterateAsync(cancellationToken);
        }

        progress?.Report("Création des thématiques à partir de vos messages…");
        var analysis = ThemeDiscoveryEngine.Discover(snapshots, profile.Account, _options.Organizer);
        return new MailboxAnalysis
        {
            DisplayName = profile.DisplayName,
            Account = profile.Account,
            AnalyzedCount = analysis.AnalyzedCount,
            ClassifiableCount = analysis.ClassifiableCount,
            UnclassifiedCount = analysis.UnclassifiedCount,
            Themes = analysis.Themes
        };
    }

    public async Task<int> ApplyThemesAsync(
        MailboxAnalysis analysis,
        IProgress<string>? progress = null,
        CancellationToken cancellationToken = default)
    {
        if (analysis.Themes.Count == 0)
        {
            return 0;
        }

        await EnsureMasterCategoriesAsync(analysis.Themes, cancellationToken);

        var applied = 0;
        var total = analysis.ClassifiableCount;
        foreach (var theme in analysis.Themes)
        {
            var name = ThemeDiscoveryEngine.SanitizeCategoryName(theme.Name);
            foreach (var messageId in theme.MessageIds)
            {
                cancellationToken.ThrowIfCancellationRequested();
                try
                {
                    await _graph.Me.Messages[messageId].PatchAsync(
                        new Message { Categories = [name] },
                        cancellationToken: cancellationToken);
                    applied++;
                    if (applied % 10 == 0 || applied == total)
                    {
                        progress?.Report($"Classement : {applied:N0} / {total:N0} messages…");
                    }
                }
                catch (Exception)
                {
                    // Un message refusé ne doit pas arrêter le reste de la boîte.
                }
            }
        }

        return applied;
    }

    private async Task EnsureMasterCategoriesAsync(
        IReadOnlyList<DiscoveredTheme> themes,
        CancellationToken cancellationToken)
    {
        var existing = await _graph.Me.Outlook.MasterCategories.GetAsync(cancellationToken: cancellationToken);
        var names = new HashSet<string>(
            existing?.Value?.Select(category => category.DisplayName ?? string.Empty) ?? [],
            StringComparer.OrdinalIgnoreCase);

        var colorIndex = 0;
        foreach (var theme in themes)
        {
            var name = ThemeDiscoveryEngine.SanitizeCategoryName(theme.Name);
            if (!names.Add(name))
            {
                continue;
            }

            await _graph.Me.Outlook.MasterCategories.PostAsync(
                new OutlookCategory
                {
                    DisplayName = name,
                    Color = PresetColors[colorIndex % PresetColors.Length]
                },
                cancellationToken: cancellationToken);

            colorIndex++;
        }
    }

    private static readonly CategoryColor[] PresetColors =
    [
        CategoryColor.Preset0, CategoryColor.Preset1, CategoryColor.Preset2, CategoryColor.Preset3,
        CategoryColor.Preset4, CategoryColor.Preset5, CategoryColor.Preset6, CategoryColor.Preset7,
        CategoryColor.Preset8, CategoryColor.Preset9, CategoryColor.Preset10, CategoryColor.Preset11,
        CategoryColor.Preset12
    ];

    private static MailSnapshot ToSnapshot(Message message)
    {
        var email = message.From?.EmailAddress?.Address ?? string.Empty;
        var at = email.IndexOf('@');
        return new MailSnapshot
        {
            Id = message.Id!,
            Subject = message.Subject ?? string.Empty,
            SenderEmail = email,
            SenderName = message.From?.EmailAddress?.Name ?? string.Empty,
            Domain = at >= 0 ? email[(at + 1)..].ToLowerInvariant() : string.Empty,
            ExistingCategories = message.Categories ?? []
        };
    }

    private static string? FirstNonEmpty(params string?[] values) =>
        values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
}
