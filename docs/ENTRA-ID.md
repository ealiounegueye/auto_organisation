# Enregistrer Outlook Organizer dans Microsoft Entra ID

Cette étape est à faire **une seule fois**, par un administrateur Microsoft 365 / Entra. Ensuite, le même Client ID est compilé dans l’application pour tous les collaborateurs.

L’application est un **client public** (logiciel installé sur le PC). Elle utilise des **permissions déléguées** : elle n’agit que sur la boîte de la personne connectée.

## 1. Ouvrir Entra

1. Allez sur [https://entra.microsoft.com](https://entra.microsoft.com) (ou Centre d’administration Microsoft 365 → Identité).
2. **Identité** → **Applications** → **Inscriptions d'applications**.
3. **Nouvelle inscription**.

## 2. Créer l’inscription

| Champ | Valeur |
|---|---|
| Nom | `Outlook Organizer` |
| Types de comptes pris en charge | **Comptes dans l’annuaire organisationnel uniquement** (locataire unique) |
| URI de redirection | **Application de bureau et mobile** → `http://localhost` |

Validez.

Copiez :

- **ID d’application (client)** → `ClientId` dans `appsettings.local.json`
- **ID de l’annuaire (locataire)** → `TenantId` dans `appsettings.local.json`

## 3. URI de redirection

**Authentification** → **Ajouter une plateforme** → **Applications mobiles et de bureau** si ce n’est pas déjà fait.

Ajoutez :

- `http://localhost`
- `ms-appx-web://microsoft.aad.brokerplugin/{VOTRE-CLIENT-ID}`

Remplacez `{VOTRE-CLIENT-ID}` par l’ID d’application.

Décochez **Jetons d’accès** / **Jetons d’ID** de flux implicite : inutiles pour cette appli de bureau.

## 4. Permissions Microsoft Graph

**Autorisations d’API** → **Ajouter une autorisation** → **Microsoft Graph** → **Autorisations déléguées** :

- `User.Read` (souvent déjà présente)
- `Mail.ReadWrite`
- `MailboxSettings.ReadWrite`

Puis **Accorder le consentement administrateur pour [votre organisation]**.

Sans ce consentement, chaque collaborateur verra une demande, ou la connexion échouera selon votre politique.

## 5. Client public

**Authentification** (ou **Paramètres avancés**) : **Autoriser les flux de clients publics** = **Oui**.

C’est nécessaire pour une application de bureau sans secret.

Ne créez **pas** de secret client (secret, certificat). Un client public n’en a pas besoin, et il ne faudrait pas le mettre dans un `.exe` distribué.

## 6. Fichier de config local

Dans le projet :

```text
OutlookOrganizer/appsettings.local.json.example
        ↓ copier
OutlookOrganizer/appsettings.local.json
```

```json
{
  "AzureAd": {
    "ClientId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
    "TenantId": "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy"
  }
}
```

`appsettings.local.json` est ignoré par git. Ne le commitez pas.

## 7. Tester

1. Ouvrez `OutlookOrganizer.sln` dans Visual Studio.
2. F5.
3. **Se connecter** avec un compte du locataire.
4. Acceptez les permissions si demandé.
5. **Analyser ma boîte**.

Résultat attendu : le compte s’affiche, un nombre de messages apparaît, le bandeau indique que **rien n’a été modifié**.

## Politique d’entreprise

Si la connexion échoue avec une erreur de type `AADSTS65001`, le consentement admin manque.

Si elle échoue avec `AADSTS700016`, le Client ID est faux ou l’application n’est pas dans ce locataire.

Les accès **application** (application permissions / client credentials) ne sont pas utilisés ici : Outlook Organizer ne doit pas lire toutes les boîtes de l’entreprise.
