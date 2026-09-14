namespace OutlookOrganizer.Models;

public static class ThemePalette
{
    public static IReadOnlyList<string> Colors { get; } =
    [
        "#0F6CBD",
        "#0F7B0F",
        "#CA5010",
        "#5C2E91",
        "#038387",
        "#C239B3",
        "#D13438",
        "#00B7C3",
        "#8B6914",
        "#107C10",
        "#881798",
        "#4F6BED",
        "#B10E1C",
        "#498205",
        "#986F0B"
    ];

    public static string At(int index) => Colors[Math.Abs(index) % Colors.Count];
}
