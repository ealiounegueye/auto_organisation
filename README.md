# Outlook Organizer

Logiciel Windows pour organiser une boîte Microsoft 365 **en un clic**.

Le collaborateur :

1. Double-clique sur l’application
2. Se connecte avec son compte Microsoft 365 (la première fois)
3. Clique sur **Organiser ma boîte**

L’application **lit ses messages** et **crée des thématiques à partir de ce qu’elle trouve** : domaines d’expéditeurs, boîtes internes du type `rh@` / `facturation@`, mots qui reviennent dans les objets. Il n’y a pas de liste imposée Finance / RH / IT.

Les catégories apparaissent ensuite dans Outlook.

## Ce que fait un clic

- Connexion Microsoft 365 si besoin (MFA compris). Le mot de passe n’est jamais stocké.
- Lecture des messages (jusqu’à 2000 par défaut).
- Découverte des thématiques **propres à cette boîte**.
- En mode réel : création des catégories Outlook et classement des messages.
- Case **Aperçu seulement** : montre les thématiques **sans rien modifier**.

Les messages déjà catégorisés, ou trop isolés, sont laissés tels quels.

## Lancer le projet (PC de test / admin)

1. Visual Studio 2022 avec la charge **Développement .NET Desktop**
2. Enregistrement Entra ID **une seule fois** : [docs/ENTRA-ID.md](docs/ENTRA-ID.md)
3. Copier `OutlookOrganizer/appsettings.local.json.example` vers `appsettings.local.json` et coller Client ID + Tenant ID
4. Ouvrir `OutlookOrganizer.sln` → **F5**

Pour un premier test, cochez **Aperçu seulement**, cliquez sur **Aperçu de ma boîte**, vérifiez les thématiques, puis décochez et cliquez sur **Organiser ma boîte**.

## Collaborateurs

Ils n’ont pas à ouvrir Visual Studio ni PowerShell. L’administrateur configure Entra + le Client ID, compile, et distribue `OutlookOrganizer.exe` (étape packaging).

Jusque-là, ne donnez pas le code source : donnez uniquement l’exécutable une fois généré.

## Permissions Graph (déléguées)

| Permission | Usage |
|---|---|
| `User.Read` | Afficher le compte |
| `Mail.ReadWrite` | Lire les messages et poser les catégories |
| `MailboxSettings.ReadWrite` | Créer les thématiques Outlook |

L’application agit **au nom de la personne connectée**, pas sur toutes les boîtes.
