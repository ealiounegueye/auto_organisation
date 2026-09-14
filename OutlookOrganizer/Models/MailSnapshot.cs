namespace OutlookOrganizer.Models;

public sealed class MailSnapshot
{
    public required string Id { get; init; }
    public string Subject { get; init; } = string.Empty;
    public string SenderEmail { get; init; } = string.Empty;
    public string SenderName { get; init; } = string.Empty;
    public string Domain { get; init; } = string.Empty;
    public IReadOnlyList<string> ExistingCategories { get; init; } = [];
}
