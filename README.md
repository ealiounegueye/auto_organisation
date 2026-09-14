# Outlook Organizer — Étape 1

Petit logiciel Windows pour organiser une boîte Microsoft 365 **sans PowerShell, sans CMD**.

Le collaborateur fera plus tard :

**double-clic → Se connecter → Analyser ma boîte → Organiser**

Cette étape 1 livre le projet WPF prêt à ouvrir dans Visual Studio, avec :

- authentification Microsoft 365 (MSAL, permissions déléguées)
- interface graphique
- **mode simulation** : l’analyse lit la boîte, elle n’écrit rien

L’étape 2 ajoutera le moteur de classement (expéditeur, objet, domaine, mots-clés, pièces jointes).  
L’étape 3 produira un seul `OutlookOrganizer.exe` distribuable.

## Ce que voit l’utilisateur

- Se connecter / Déconnexion
- Mode simulation (activé par défaut)
- Analyser ma boîte (lecture seule via Microsoft Graph)
- Compteurs par catégorie (Finance, RH, IT, Achats, Projets, Newsletters, Autres)
- Organiser — en étape 1, confirme la simulation et **ne modifie aucun message**

Le mot de passe n’est **jamais** demandé par l’application ni stocké. Microsoft 365 gère la connexion (et le MFA s’il est activé).

## Prérequis

Sur le PC de développement (Windows) :

1. [Visual Studio 2022](https://visualstudio.microsoft.com/) avec la charge **Développement .NET Desktop**
2. [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0) (inclus avec VS si la charge est cochée)
3. Une application enregistrée dans **Microsoft Entra ID** — voir [docs/ENTRA-ID.md](docs/ENTRA-ID.md)

Le projet est du WPF Windows : il se compile sur un PC Windows, pas sur Mac.

## Démarrage rapide

1. Enregistrez l’application dans Entra ID et copiez le **Client ID**.
2. Copiez `OutlookOrganizer/appsettings.local.json.example` vers `OutlookOrganizer/appsettings.local.json`.
3. Remplacez les GUID :

```json
{
  "AzureAd": {
    "ClientId": "votre-client-id",
    "TenantId": "votre-tenant-id"
  }
}
```

`TenantId` : l’identifiant du locataire Entra (recommandé en entreprise).  
Vous pouvez laisser `"organizations"` le temps des tests.

4. Ouvrez `OutlookOrganizer.sln` dans Visual Studio.
5. F5 pour lancer.

Au premier clic sur **Se connecter**, Microsoft affiche la fenêtre de connexion habituelle.

## Permissions Graph (déléguées)

L’application agit **au nom de l’utilisateur connecté**, pas sur toutes les boîtes de l’entreprise.

| Permission                 | Usage                                      |
|----------------------------|--------------------------------------------|
| `User.Read`                | Afficher le compte connecté                |
| `Mail.ReadWrite`           | Lire les messages ; écrire les catégories (étape 2) |
| `MailboxSettings.ReadWrite`| Créer les catégories Outlook (étape 2)     |

En étape 1, seul la **lecture** est réellement utilisée. Les droits d’écriture sont déjà demandés pour ne pas changer le consentement plus tard.

L’administrateur peut devoir **accorder le consentement admin**.

## Structure

```
OutlookOrganizer.sln
OutlookOrganizer/
  App.xaml
  MainWindow.xaml          Interface
  appsettings.json         Config par défaut
  Configuration/           Client ID, limites d'analyse
  Services/                MSAL + Microsoft Graph
  ViewModels/              Boutons et états de l'écran
  Models/                  Catégories et résultat d'analyse
docs/ENTRA-ID.md           Enregistrement Entra, étape par étape
```

## Distribution

Pas encore. L’étape 3 produira un `.exe` unique (Teams, SharePoint, Intune, clé USB).

Pour l’instant, ne donnez **pas** ce projet aux collaborateurs.
