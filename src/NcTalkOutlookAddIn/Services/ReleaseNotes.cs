// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    /// <summary>
    /// User-facing notes for one version: what changed, what the user has to do, what they gain.
    /// </summary>
    internal sealed class ReleaseNote
    {
        internal string Version { get; set; }
        internal string[] WhatsNew { get; set; }
        internal string[] ToDo { get; set; }
        internal string[] Benefits { get; set; }
    }

    /// <summary>
    /// "What's new" after an update. Shown once per version that has notes; versions without notes
    /// (small fixes) stay silent. The last shown version is kept in whats-new.xml.
    /// </summary>
    internal static class ReleaseNotes
    {
        private const string StateFileName = "whats-new.xml";

        private static readonly ReleaseNote[] Notes =
        {
            new ReleaseNote
            {
                Version = "3.5.0",
                WhatsNew = new[]
                {
                    "Ihr Nextcloud-Kalender wird automatisch mit Ihrem Outlook-Kalender abgeglichen – in beide Richtungen.",
                    "Der Outlook-Ordner \"Kunden\" wird mit Ihrem Nextcloud-Adressbuch \"IBP-Kunden\" abgeglichen, inklusive Kontaktfotos.",
                    "Ihre Kontakte kommen jetzt gezielt aus dem Adressbuch \"IBP-Kontakte\".",
                    "Doppelte Gruppenordner (z. B. \"Sifa\" zweimal) werden automatisch aufgeräumt.",
                    "In den Einstellungen (Reiter \"Kontakte\") gibt es den Knopf \"Kontakte-Sync zurücksetzen und neu aufbauen\"."
                },
                ToDo = new[]
                {
                    "Wenn der Hinweis zum Kalender erscheint: Outlook einmal neu starten. Erst danach werden Ihre Termine abgeglichen.",
                    "Vorher legt NC Connector eine Sicherungskopie Ihres Kalenders an (Ordner \"Kalender - Sicherung vor Nextcloud-Sync …\"). "
                        + "Wenn nach ein paar Tagen alle Termine stimmen, können Sie diesen Ordner löschen.",
                    "Neue Kundenkontakte legen Sie im Outlook-Ordner \"Kunden\" an.",
                    "Sonst müssen Sie nichts tun – alles Weitere läuft automatisch."
                },
                Benefits = new[]
                {
                    "Ihre Termine sind überall aktuell: in Outlook, in Nextcloud im Browser und auf dem Smartphone.",
                    "Ihre Kundenkontakte stehen Ihnen auf allen Geräten zur Verfügung.",
                    "Keine doppelten Gruppenordner mehr in Ihren Kontakten.",
                    "Stimmt bei den Kontakten einmal etwas nicht, baut ein Klick alles sauber neu auf."
                }
            }
        };

        /// <summary>
        /// Notes to show now, or null when this version was already shown or has no notes.
        /// </summary>
        internal static ReleaseNote Pending(string currentVersion)
        {
            string version = ShortVersion(currentVersion);
            ReleaseNote note = Notes.FirstOrDefault(n => string.Equals(n.Version, version, StringComparison.Ordinal));
            if (note == null)
            {
                return null;
            }
            return string.Equals(LoadLastShownVersion(), version, StringComparison.Ordinal) ? null : note;
        }

        internal static void MarkShown(ReleaseNote note)
        {
            if (note == null) return;
            try
            {
                new XDocument(new XElement("WhatsNew", new XAttribute("lastShownVersion", note.Version)))
                    .Save(GetStatePath());
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to save what's-new state.", ex);
            }
        }

        internal static string ShortVersion(string version)
        {
            Version parsed;
            return System.Version.TryParse(version ?? string.Empty, out parsed)
                ? parsed.Major + "." + parsed.Minor + "." + Math.Max(0, parsed.Build)
                : (version ?? string.Empty).Trim();
        }

        private static string LoadLastShownVersion()
        {
            try
            {
                string path = GetStatePath();
                if (!File.Exists(path)) return string.Empty;
                XElement root = XDocument.Load(path).Root;
                return root != null ? ((string)root.Attribute("lastShownVersion") ?? string.Empty) : string.Empty;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to load what's-new state.", ex);
                return string.Empty;
            }
        }

        private static string GetStatePath()
        {
            return Path.Combine(AppDataPaths.EnsureLocalRootDirectory(), StateFileName);
        }
    }
}
