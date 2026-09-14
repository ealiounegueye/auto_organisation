# Outlook Organizer

Logiciel Windows pour organiser une boîte Microsoft 365 **en un clic**.

Les collaborateurs n’ont **pas** besoin de Visual Studio.

1. Double-clic sur `OutlookOrganizer.exe`
2. Connexion Microsoft 365 (la première fois)
3. **Organiser ma boîte**

L’application lit les messages et **crée les thématiques à partir de cette boîte** (domaines, boîtes internes, mots des objets). Rien n’est imposé à l’avance.

## Deux rôles

| Qui | Ce qu’il a sur son PC |
|---|---|
| Collaborateur | Rien. Juste le fichier `.exe` |
| Vous (admin / test) | Le [SDK .NET 8](https://dotnet.microsoft.com/download/dotnet/8.0) **une fois**, pour fabriquer l’exe. Visual Studio est facultatif |

## 1. Préparer Entra ID (une seule fois)

Suivez [docs/ENTRA-ID.md](docs/ENTRA-ID.md), puis dans `OutlookOrganizer/appsettings.json` mettez :

- `ClientId` : ID d’application
- `TenantId` : ID du locataire (ou `organizations` le temps du test)

Sans ça, l’appli s’ouvre mais la connexion Microsoft 365 échoue.

## 2. Fabriquer l’exe (un PC Windows admin)

1. Installez le **SDK .NET 8** (pas Visual Studio) : https://dotnet.microsoft.com/download/dotnet/8.0  
   Choisissez le SDK 8.0 Windows x64.
2. Récupérez le projet :

```powershell
git clone https://github.com/ealiounegueye/auto_organisation.git
cd auto_organisation
git pull
```

3. Double-cliquez sur **`Publier.bat`**.
4. Le dossier `dist` contient `OutlookOrganizer.exe` + `appsettings.json`.

C’est ce dossier que vous copiez (Teams, SharePoint, clé USB). Pas le code source.

## 3. Comment tester

Sur **votre** PC Windows, dans `dist` :

1. Double-clic sur `OutlookOrganizer.exe`.
2. Cochez **Aperçu seulement**.
3. Cliquez sur **Aperçu de ma boîte** (ou Organiser : l’appli demandera la connexion Microsoft 365).
4. Vérifiez les thématiques proposées. Rien n’est écrit dans Outlook.
5. Décochez l’aperçu, cliquez sur **Organiser ma boîte**.
6. Ouvrez Outlook : les catégories doivent apparaître sur les messages classés.

Si la fenêtre Microsoft 365 ne s’ouvre pas : Client ID / consentement admin (voir `docs/ENTRA-ID.md`).

Testez d’abord sur **votre** boîte, puis sur 1 ou 2 collègues, avant de le donner à tout le monde.

## Collaborateurs

Ils reçoivent `OutlookOrganizer.exe` (+ `appsettings.json` à côté).  
Double-clic. Pas de Visual Studio, pas de .NET, pas de CMD.

## Permissions Graph (déléguées)

| Permission | Usage |
|---|---|
| `User.Read` | Afficher le compte |
| `Mail.ReadWrite` | Lire les messages et poser les catégories |
| `MailboxSettings.ReadWrite` | Créer les thématiques Outlook |

L’application agit **au nom de la personne connectée**, pas sur toutes les boîtes.
