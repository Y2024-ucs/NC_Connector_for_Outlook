// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class CalDavCalendarSyncState
    {
        internal string CalendarHref { get; set; }
        internal bool BaselineEstablished { get; set; }
        internal DateTime? BaselineEstablishedUtc { get; set; }
    }

    internal sealed class CalDavSyncStateStore
    {
        private const string FileName = "caldav-sync-state.xml";
        private readonly Dictionary<string, CalDavCalendarSyncState> _states =
            new Dictionary<string, CalDavCalendarSyncState>(StringComparer.OrdinalIgnoreCase);

        internal static CalDavSyncStateStore Load()
        {
            var store = new CalDavSyncStateStore();
            string path = GetPath();
            if (!File.Exists(path))
            {
                return store;
            }

            try
            {
                var document = new XmlDocument();
                document.Load(path);
                XmlElement root = document.DocumentElement;
                if (root == null || !string.Equals(root.Name, "CalDavSyncState", StringComparison.OrdinalIgnoreCase))
                {
                    return store;
                }

                foreach (XmlNode node in root.SelectNodes("Calendar"))
                {
                    XmlElement element = node as XmlElement;
                    if (element == null)
                    {
                        continue;
                    }
                    string href = (element.GetAttribute("href") ?? string.Empty).Trim();
                    if (href.Length == 0)
                    {
                        continue;
                    }

                    bool baseline;
                    DateTime parsedUtc;
                    string utcText = element.GetAttribute("baselineEstablishedUtc") ?? string.Empty;
                    var state = new CalDavCalendarSyncState
                    {
                        CalendarHref = href,
                        BaselineEstablished = bool.TryParse(element.GetAttribute("baselineEstablished"), out baseline) && baseline,
                        BaselineEstablishedUtc = DateTime.TryParse(
                            utcText,
                            System.Globalization.CultureInfo.InvariantCulture,
                            System.Globalization.DateTimeStyles.RoundtripKind,
                            out parsedUtc)
                            ? (DateTime?)parsedUtc.ToUniversalTime()
                            : null
                    };
                    store._states[href] = state;
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to load CalDAV sync state.", ex);
            }

            return store;
        }

        internal CalDavCalendarSyncState Get(string calendarHref)
        {
            string href = (calendarHref ?? string.Empty).Trim();
            CalDavCalendarSyncState state;
            if (href.Length > 0 && _states.TryGetValue(href, out state))
            {
                return state;
            }
            return new CalDavCalendarSyncState
            {
                CalendarHref = href,
                BaselineEstablished = false,
                BaselineEstablishedUtc = null
            };
        }

        internal bool CanDelete(string calendarHref)
        {
            return Get(calendarHref).BaselineEstablished;
        }

        internal void MarkBaselineEstablished(string calendarHref)
        {
            string href = (calendarHref ?? string.Empty).Trim();
            if (href.Length == 0)
            {
                return;
            }
            _states[href] = new CalDavCalendarSyncState
            {
                CalendarHref = href,
                BaselineEstablished = true,
                BaselineEstablishedUtc = DateTime.UtcNow
            };
        }

        internal void Save()
        {
            string path = GetPath();
            try
            {
                var document = new XmlDocument();
                XmlElement root = document.CreateElement("CalDavSyncState");
                document.AppendChild(root);

                foreach (CalDavCalendarSyncState state in _states.Values)
                {
                    XmlElement calendar = document.CreateElement("Calendar");
                    calendar.SetAttribute("href", state.CalendarHref ?? string.Empty);
                    calendar.SetAttribute("baselineEstablished", state.BaselineEstablished.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    if (state.BaselineEstablishedUtc.HasValue)
                    {
                        calendar.SetAttribute(
                            "baselineEstablishedUtc",
                            state.BaselineEstablishedUtc.Value.ToUniversalTime().ToString("o", System.Globalization.CultureInfo.InvariantCulture));
                    }
                    root.AppendChild(calendar);
                }

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
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to save CalDAV sync state.", ex);
                throw;
            }
        }

        private static string GetPath()
        {
            return Path.Combine(AppDataPaths.EnsureLocalRootDirectory(), FileName);
        }
    }
}
