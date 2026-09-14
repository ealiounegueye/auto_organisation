# Outlook Organizer

Application desktop Windows. Vous envoyez **un fichier**. Le collaborateur **double-clique**. Le reste se fait tout seul.

1. Fenêtre Microsoft 365 (connexion, une fois)
2. Lecture de sa boîte
3. Création des thématiques à partir de **ses** messages
4. Classement dans Outlook
5. Message « C’est terminé »

Pas de Visual Studio, pas de CMD, pas de bouton à comprendre.

## Vous (une fois)

1. Enregistrez l’appli dans Entra ID : [docs/ENTRA-ID.md](docs/ENTRA-ID.md)
2. Collez le Client ID dans `OutlookOrganizer/appsettings.json`
3. Sur un PC Windows, installez le [SDK .NET 8](https://dotnet.microsoft.com/download/dotnet/8.0)
4. Double-clic sur `Publier.bat`
5. Envoyez `dist\OutlookOrganizer.exe` (Teams, mail, USB)

## Eux

Double-clic sur `OutlookOrganizer.exe`. C’est tout.

## Tester chez vous

Même chose : lancez l’exe, connectez-vous avec **votre** compte, laissez tourner, ouvrez Outlook.

Pour un essai **sans** modifier Outlook :

```text
OutlookOrganizer.exe --preview
```
