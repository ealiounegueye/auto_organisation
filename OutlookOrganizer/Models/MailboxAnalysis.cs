namespace OutlookOrganizer.Models;

public sealed class MailboxAnalysis
{
    public string? DisplayName { get; init; }
    public string? Account { get; init; }
    public int AnalyzedCount { get; init; }
    public int ClassifiableCount { get; init; }
    public IReadOnlyDictionary<string, int> CategoryCounts { get; init; } =
        new Dictionary<string, int>();
}
