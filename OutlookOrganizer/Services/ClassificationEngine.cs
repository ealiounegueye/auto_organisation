using Microsoft.Graph.Models;

namespace OutlookOrganizer.Services;

/// <summary>
/// Étape 2 : classer un message à partir de l'expéditeur, de l'objet,
/// du domaine, des mots-clés et des pièces jointes.
/// En étape 1, aucun message n'est classé — ils restent dans Autres.
/// </summary>
public static class ClassificationEngine
{
    public static string? Classify(Message message)
    {
        _ = message;
        return null;
    }
}
