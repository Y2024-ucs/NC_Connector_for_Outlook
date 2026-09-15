// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    internal enum CalDavSyncDirection
    {
        TwoWay = 0,
        NextcloudToOutlook = 1,
        OutlookToNextcloud = 2
    }

    internal sealed class CalDavSyncPreferences
    {
        private const string FileName = "caldav-sync.xml";

        internal CalDavSyncPreferences()
        {
            Configured = false;
            Enabled = true;
            Direction = CalDavSyncDirection.TwoWay;
            UseDefaultCalendarFolder = false;
            SyncOnStartup = true;
            IntervalMinutes = 15;
            SelectedCalendarHrefs = new List<string>();
        }

        internal bool Configured { get; set; }
        internal bool Enabled { get; set; }
        internal CalDavSyncDirection Direction { get; set; }
        internal bool UseDefaultCalendarFolder { get; set; }
        internal bool SyncOnStartup { get; set; }
        internal int IntervalMinutes { get; set; }
        internal List<string> SelectedCalendarHrefs { get; private set; }

        internal bool IsCalendarSelected(string href)
        {
            return !string.IsNullOrWhiteSpace(href)
                && SelectedCalendarHrefs.Any(v => string.Equals(v, href, StringComparison.OrdinalIgnoreCase));
        }

        internal void SetSelectedCalendars(IEnumerable<string> hrefs)
        {
            SelectedCalendarHrefs.Clear();
            foreach (string href in hrefs ?? Enumerable.Empty<string>())
            {
                string value = (href ?? string.Empty).Trim();
                if (value.Length == 0 || IsCalendarSelected(value))
                {
                    continue;
                }
                SelectedCalendarHrefs.Add(value);
            }
        }

        internal static CalDavSyncPreferences Load()
        {
            var preferences = new CalDavSyncPreferences();
            string path = GetPath();
            if (!File.Exists(path))
            {
                return preferences;
            }

            try
            {
                var document = new XmlDocument();
                document.Load(path);
                XmlElement root = document.DocumentElement;
                if (root == null || !string.Equals(root.Name, "CalDavSync", StringComparison.OrdinalIgnoreCase))
                {
                    return preferences;
                }

                preferences.Configured = ReadBool(root, "Configured", false);
                preferences.Enabled = ReadBool(root, "Enabled", true);
                preferences.UseDefaultCalendarFolder = ReadBool(root, "UseDefaultCalendarFolder", false);
                preferences.SyncOnStartup = ReadBool(root, "SyncOnStartup", true);
                preferences.IntervalMinutes = Math.Max(1, ReadInt(root, "IntervalMinutes", 15));

                CalDavSyncDirection direction;
                string directionText = ReadText(root, "Direction");
                preferences.Direction = Enum.TryParse(directionText, true, out direction)
                    ? direction
                    : CalDavSyncDirection.TwoWay;

                XmlElement selected = root["SelectedCalendars"];
                if (selected != null)
                {
                    foreach (XmlNode node in selected.SelectNodes("Calendar"))
                    {
                        XmlElement element = node as XmlElement;
                        if (element == null)
                        {
                            continue;
                        }
                        string href = (element.GetAttribute("href") ?? string.Empty).Trim();
                        if (href.Length > 0 && !preferences.IsCalendarSelected(href))
                        {
                            preferences.SelectedCalendarHrefs.Add(href);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to load CalDAV sync preferences.", ex);
            }
            return preferences;
        }

        internal void Save()
        {
            string path = GetPath();
            try
            {
                var document = new XmlDocument();
                XmlElement root = document.CreateElement("CalDavSync");
                document.AppendChild(root);
                AppendBool(document, root, "Configured", Configured);
                AppendBool(document, root, "Enabled", Enabled);
                AppendText(document, root, "Direction", Direction.ToString());
                AppendBool(document, root, "UseDefaultCalendarFolder", UseDefaultCalendarFolder);
                AppendBool(document, root, "SyncOnStartup", SyncOnStartup);
                AppendText(document, root, "IntervalMinutes", Math.Max(1, IntervalMinutes).ToString(System.Globalization.CultureInfo.InvariantCulture));

                XmlElement selected = document.CreateElement("SelectedCalendars");
                foreach (string href in SelectedCalendarHrefs)
                {
                    XmlElement calendar = document.CreateElement("Calendar");
                    calendar.SetAttribute("href", href ?? string.Empty);
                    selected.AppendChild(calendar);
                }
                root.AppendChild(selected);

                var settings = new XmlWriterSettings
                {
                    Indent = true,
                    Encoding = new System.Text.UTF8Encoding(false)
                };
                using (XmlWriter writer = XmlWriter.Create(path, settings))
                {
                    document.Save(writer);
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to save CalDAV sync preferences.", ex);
                throw;
            }
        }

        private static string GetPath()
        {
            return Path.Combine(AppDataPaths.EnsureLocalRootDirectory(), FileName);
        }

        private static bool ReadBool(XmlElement root, string name, bool fallback)
        {
            bool value;
            string text = ReadText(root, name);
            return bool.TryParse(text, out value) ? value : fallback;
        }

        private static int ReadInt(XmlElement root, string name, int fallback)
        {
            int value;
            string text = ReadText(root, name);
            return int.TryParse(text, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out value)
                ? value
                : fallback;
        }

        private static string ReadText(XmlElement root, string name)
        {
            XmlElement element = root[name];
            return element != null ? element.InnerText ?? string.Empty : string.Empty;
        }

        private static void AppendBool(XmlDocument document, XmlElement root, string name, bool value)
        {
            AppendText(document, root, name, value.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        private static void AppendText(XmlDocument document, XmlElement root, string name, string value)
        {
            XmlElement element = document.CreateElement(name);
            element.InnerText = value ?? string.Empty;
            root.AppendChild(element);
        }
    }
}
