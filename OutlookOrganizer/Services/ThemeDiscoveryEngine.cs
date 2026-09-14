using System.Globalization;
using System.Text.RegularExpressions;
using OutlookOrganizer.Configuration;
using OutlookOrganizer.Models;

namespace OutlookOrganizer.Services;

/// <summary>
/// Découvre des thématiques à partir du contenu réel de la boîte :
/// domaines d'expéditeurs, boîtes fonctionnelles internes, mots des objets.
/// Aucune liste Finance/RH/IT n'est imposée à l'avance.
/// </summary>
public static class ThemeDiscoveryEngine
{
    private static readonly HashSet<string> PublicDomains = new(StringComparer.OrdinalIgnoreCase)
    {
        "gmail.com", "googlemail.com", "outlook.com", "outlook.fr", "hotmail.com", "hotmail.fr",
        "live.com", "live.fr", "msn.com", "yahoo.com", "yahoo.fr", "icloud.com", "me.com",
        "aol.com", "proton.me", "protonmail.com", "orange.fr", "wanadoo.fr", "free.fr",
        "sfr.fr", "laposte.net", "bbox.fr"
    };

    private static readonly HashSet<string> StopWords = new(StringComparer.OrdinalIgnoreCase)
    {
        "le", "la", "les", "un", "une", "des", "de", "du", "et", "ou", "au", "aux", "en", "dans",
        "pour", "par", "sur", "pas", "ne", "que", "qui", "dont", "est", "sont", "avec", "sans",
        "plus", "moins", "très", "comme", "mais", "donc", "alors", "cette", "cet", "ces", "mon",
        "ma", "mes", "ton", "ta", "tes", "son", "sa", "ses", "notre", "nos", "votre", "vos",
        "leur", "leurs", "the", "and", "for", "from", "you", "your", "this", "that", "with",
        "object", "objet", "mail", "email", "message", "re", "tr", "fw", "fwd", "rv", "https",
        "http", "www", "com", "fr", "sn", "info", "please", "merci", "bonjour", "cordialement",
        "equipe", "équipe", "hello", "thanks", "request", "notification"
    };

    private static readonly Regex NewsletterSender = new(
        @"noreply|no-reply|donotreply|do-not-reply|newsletter|notifications?|mailer|daemon|bounce",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex SubjectPrefix = new(
        @"^((re|tr|fw|fwd|rv|ref)\s*:|\s*\[(externe|external|external mail)\]\s*)+",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

    public static MailboxAnalysis Discover(
        IReadOnlyList<MailSnapshot> messages,
        string? userAccount,
        OrganizerOptions options)
    {
        var unclassifiedMessages = messages
            .Where(message => message.ExistingCategories.Count == 0)
            .ToList();
        var userDomain = DomainOf(userAccount);
        var minCount = Math.Max(4, options.MinEmailsPerTheme);
        var maxThemes = Math.Max(3, options.MaxThemes);
        var assigned = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var themes = new List<DiscoveredTheme>();

        AddDomainThemes(unclassifiedMessages, userDomain, minCount, assigned, themes);
        AddFunctionalMailboxThemes(unclassifiedMessages, userDomain, minCount, assigned, themes);
        AddNewsletterTheme(unclassifiedMessages, minCount, assigned, themes);
        AddSubjectKeywordThemes(unclassifiedMessages, minCount, maxThemes, assigned, themes);

        var ordered = themes
            .Where(theme => theme.Count > 0)
            .OrderByDescending(theme => theme.Count)
            .Take(maxThemes)
            .Select((theme, index) => new DiscoveredTheme
            {
                Name = theme.Name,
                ColorHex = ThemePalette.At(index),
                Source = theme.Source,
                MessageIds = theme.MessageIds
            })
            .ToList();

        var classifiable = ordered.Sum(theme => theme.Count);
        var unclassified = Math.Max(0, unclassifiedMessages.Count - classifiable);

        return new MailboxAnalysis
        {
            AnalyzedCount = messages.Count,
            ClassifiableCount = classifiable,
            UnclassifiedCount = unclassified,
            Themes = ordered
        };
    }

    private static void AddDomainThemes(
        IReadOnlyList<MailSnapshot> messages,
        string? userDomain,
        int minCount,
        Dictionary<string, string> assigned,
        List<DiscoveredTheme> themes)
    {
        var groups = messages
            .Where(message => !assigned.ContainsKey(message.Id))
            .Select(message => (message, key: RegistrableDomain(message.Domain)))
            .Where(item =>
                !string.IsNullOrWhiteSpace(item.key)
                && !PublicDomains.Contains(item.key)
                && !IsSameOrganization(item.key, userDomain))
            .GroupBy(item => item.key, StringComparer.OrdinalIgnoreCase);

        foreach (var group in groups.OrderByDescending(item => item.Count()))
        {
            if (group.Count() < minCount)
            {
                continue;
            }

            var theme = new DiscoveredTheme
            {
                Name = DisplayNameFromDomain(group.Key),
                ColorHex = "#0F6CBD",
                Source = $"domaine {group.Key}"
            };

            foreach (var item in group)
            {
                if (assigned.TryAdd(item.message.Id, theme.Name))
                {
                    theme.MessageIds.Add(item.message.Id);
                }
            }

            themes.Add(theme);
        }
    }

    private static void AddFunctionalMailboxThemes(
        IReadOnlyList<MailSnapshot> messages,
        string? userDomain,
        int minCount,
        Dictionary<string, string> assigned,
        List<DiscoveredTheme> themes)
    {
        if (string.IsNullOrWhiteSpace(userDomain))
        {
            return;
        }

        var groups = messages
            .Where(message => !assigned.ContainsKey(message.Id) && IsSameOrganization(message.Domain, userDomain))
            .Select(message => (message, local: LocalPart(message.SenderEmail)))
            .Where(item => IsFunctionalMailbox(item.local))
            .GroupBy(item => item.local, StringComparer.OrdinalIgnoreCase);

        foreach (var group in groups.OrderByDescending(item => item.Count()))
        {
            if (group.Count() < minCount)
            {
                continue;
            }

            var theme = new DiscoveredTheme
            {
                Name = TitleCase(group.Key.Replace('-', ' ').Replace('_', ' ')),
                ColorHex = "#0F6CBD",
                Source = $"boîte {group.Key}@{userDomain}"
            };

            foreach (var item in group)
            {
                if (assigned.TryAdd(item.message.Id, theme.Name))
                {
                    theme.MessageIds.Add(item.message.Id);
                }
            }

            themes.Add(theme);
        }
    }

    private static void AddNewsletterTheme(
        IReadOnlyList<MailSnapshot> messages,
        int minCount,
        Dictionary<string, string> assigned,
        List<DiscoveredTheme> themes)
    {
        var matches = messages
            .Where(message => !assigned.ContainsKey(message.Id) && LooksLikeNewsletter(message))
            .ToList();

        if (matches.Count < minCount)
        {
            return;
        }

        var theme = new DiscoveredTheme
        {
            Name = "Newsletters",
            ColorHex = "#69797E",
            Source = "expéditeurs automatiques détectés"
        };

        foreach (var message in matches)
        {
            if (assigned.TryAdd(message.Id, theme.Name))
            {
                theme.MessageIds.Add(message.Id);
            }
        }

        themes.Add(theme);
    }

    private static void AddSubjectKeywordThemes(
        IReadOnlyList<MailSnapshot> messages,
        int minCount,
        int maxThemes,
        Dictionary<string, string> assigned,
        List<DiscoveredTheme> themes)
    {
        var remaining = messages.Where(message => !assigned.ContainsKey(message.Id)).ToList();
        if (remaining.Count == 0)
        {
            return;
        }

        var keywordHits = new Dictionary<string, List<MailSnapshot>>(StringComparer.OrdinalIgnoreCase);
        foreach (var message in remaining)
        {
            foreach (var token in TokensOf(message.Subject))
            {
                if (!keywordHits.TryGetValue(token, out var list))
                {
                    list = [];
                    keywordHits[token] = list;
                }

                if (list.All(existing => existing.Id != message.Id))
                {
                    list.Add(message);
                }
            }
        }

        var remainingSlots = Math.Max(0, maxThemes - themes.Count);
        var candidates = keywordHits
            .Where(pair => pair.Value.Count >= minCount)
            .OrderByDescending(pair => pair.Value.Count)
            .ThenByDescending(pair => pair.Key.Length)
            .Take(remainingSlots);

        foreach (var candidate in candidates)
        {
            var stillFree = candidate.Value.Where(message => !assigned.ContainsKey(message.Id)).ToList();
            if (stillFree.Count < minCount)
            {
                continue;
            }

            var theme = new DiscoveredTheme
            {
                Name = TitleCase(candidate.Key),
                ColorHex = "#0F6CBD",
                Source = $"mot d'objet « {candidate.Key} »"
            };

            foreach (var message in stillFree)
            {
                if (assigned.TryAdd(message.Id, theme.Name))
                {
                    theme.MessageIds.Add(message.Id);
                }
            }

            themes.Add(theme);
        }
    }

    private static IEnumerable<string> TokensOf(string subject)
    {
        var cleaned = SubjectPrefix.Replace(subject ?? string.Empty, string.Empty);
        foreach (Match match in Regex.Matches(cleaned.ToLowerInvariant(), @"[\p{L}]{4,}"))
        {
            var token = match.Value;
            if (!StopWords.Contains(token))
            {
                yield return token;
            }
        }
    }

    private static bool LooksLikeNewsletter(MailSnapshot message) =>
        NewsletterSender.IsMatch(message.SenderEmail)
        || NewsletterSender.IsMatch(message.SenderName);

    private static bool IsFunctionalMailbox(string localPart)
    {
        if (string.IsNullOrWhiteSpace(localPart) || localPart.Length < 2)
        {
            return false;
        }

        if (localPart.Contains('.') || localPart.Contains('+'))
        {
            return false;
        }

        return !Regex.IsMatch(localPart, @"^\d+$");
    }

    private static string LocalPart(string email)
    {
        var at = email.IndexOf('@');
        return at <= 0 ? string.Empty : email[..at];
    }

    private static string DomainOf(string? email)
    {
        if (string.IsNullOrWhiteSpace(email))
        {
            return string.Empty;
        }

        var at = email.IndexOf('@');
        return at < 0 ? string.Empty : email[(at + 1)..].Trim().ToLowerInvariant();
    }

    private static string RegistrableDomain(string domain)
    {
        if (string.IsNullOrWhiteSpace(domain))
        {
            return string.Empty;
        }

        var parts = domain.Split('.', StringSplitOptions.RemoveEmptyEntries);
        return parts.Length >= 2 ? $"{parts[^2]}.{parts[^1]}" : domain;
    }

    private static bool IsSameOrganization(string domain, string? userDomain)
    {
        if (string.IsNullOrWhiteSpace(domain) || string.IsNullOrWhiteSpace(userDomain))
        {
            return false;
        }

        var left = RegistrableDomain(domain);
        var right = RegistrableDomain(userDomain);
        return left.Equals(right, StringComparison.OrdinalIgnoreCase);
    }

    private static string DisplayNameFromDomain(string domain)
    {
        var parts = domain.Split('.', StringSplitOptions.RemoveEmptyEntries);
        var raw = parts.Length == 0 ? domain : parts[0];
        return TitleCase(raw.Replace('-', ' '));
    }

    private static string TitleCase(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "Thématique";
        }

        var text = value.Trim();
        if (text.Length > 40)
        {
            text = text[..40];
        }

        return CultureInfo.GetCultureInfo("fr-FR").TextInfo.ToTitleCase(text.ToLowerInvariant());
    }

    public static string SanitizeCategoryName(string name)
    {
        var cleaned = Regex.Replace(name, @"[^\p{L}\p{N} \-_']+", " ").Trim();
        if (string.IsNullOrWhiteSpace(cleaned))
        {
            cleaned = "Thematique";
        }

        return cleaned.Length <= 255 ? cleaned : cleaned[..255];
    }
}
