namespace OutlookOrganizer.Models;

public sealed class MailboxAnalysis
{
    public string? DisplayName { get; init; }
    public string? Account { get; init; }
    public int AnalyzedCount { get; init; }
    public int ClassifiableCount { get; init; }
    public int UnclassifiedCount { get; init; }
    public IReadOnlyList<DiscoveredTheme> Themes { get; init; } = [];
}
