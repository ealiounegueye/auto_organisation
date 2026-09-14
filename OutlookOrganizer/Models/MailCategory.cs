namespace OutlookOrganizer.Models;

public sealed class MailCategory
{
    public required string Name { get; init; }
    public required string ColorKey { get; init; }

    public static IReadOnlyList<MailCategory> All { get; } =
    [
        new() { Name = "Finance", ColorKey = "#0F7B0F" },
        new() { Name = "RH", ColorKey = "#5C2E91" },
        new() { Name = "IT", ColorKey = "#0078D4" },
        new() { Name = "Achats", ColorKey = "#CA5010" },
        new() { Name = "Projets", ColorKey = "#038387" },
        new() { Name = "Newsletters", ColorKey = "#69797E" },
        new() { Name = "Autres", ColorKey = "#605E5C" }
    ];
}
