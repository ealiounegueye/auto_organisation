namespace OutlookOrganizer.Configuration;

public sealed class AzureAdOptions
{
    public string ClientId { get; set; } = string.Empty;
    public string TenantId { get; set; } = "organizations";
    public string RedirectUri { get; set; } = "http://localhost";
}

public sealed class OrganizerOptions
{
    public int MaxMessagesToAnalyze { get; set; } = 2000;
    public bool SimulationModeDefault { get; set; } = true;
}

public sealed class AppOptions
{
    public AzureAdOptions AzureAd { get; set; } = new();
    public OrganizerOptions Organizer { get; set; } = new();

    public bool HasValidClientId =>
        !string.IsNullOrWhiteSpace(AzureAd.ClientId)
        && AzureAd.ClientId != "REMPLACER-PAR-VOTRE-CLIENT-ID"
        && AzureAd.ClientId != "00000000-0000-0000-0000-000000000000";
}
