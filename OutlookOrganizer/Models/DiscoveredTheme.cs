namespace OutlookOrganizer.Models;

public sealed class DiscoveredTheme
{
    public required string Name { get; init; }
    public required string ColorHex { get; init; }
    public required string Source { get; init; }
    public List<string> MessageIds { get; init; } = [];
    public int Count => MessageIds.Count;
}
