using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions.Authentication;
using OutlookOrganizer.Configuration;
using OutlookOrganizer.Models;

namespace OutlookOrganizer.Services;

public sealed class GraphMailService
{
    private readonly GraphServiceClient _graph;
    private readonly int _maxMessages;

    public GraphMailService(GraphAuthService auth, AppOptions options)
    {
        _graph = new GraphServiceClient(new BaseBearerTokenAuthenticationProvider(new MsalTokenProvider(auth)));
        _maxMessages = Math.Max(50, options.Organizer.MaxMessagesToAnalyze);
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
        var counts = MailCategory.All.ToDictionary(category => category.Name, _ => 0);
        var analyzed = 0;
        var classifiable = 0;

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
                    analyzed++;
                    var category = ClassificationEngine.Classify(message) ?? "Autres";
                    if (!counts.ContainsKey(category))
                    {
                        category = "Autres";
                    }

                    counts[category]++;
                    if (category != "Autres")
                    {
                        classifiable++;
                    }

                    progress?.Report($"{analyzed:N0} messages lus…");
                    return analyzed < _maxMessages;
                });

            await iterator.IterateAsync(cancellationToken);
        }

        return new MailboxAnalysis
        {
            DisplayName = profile.DisplayName,
            Account = profile.Account,
            AnalyzedCount = analyzed,
            ClassifiableCount = classifiable,
            CategoryCounts = counts
        };
    }

    private static string? FirstNonEmpty(params string?[] values) =>
        values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
}
