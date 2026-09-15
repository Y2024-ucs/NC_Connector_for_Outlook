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

    internal sealed class CalDavItemSyncState
    {
        internal string CalendarHref { get; set; }
        internal string Uid { get; set; }
        internal string Href { get; set; }
        internal string ETag { get; set; }
        internal DateTime? LastSeenBothUtc { get; set; }
    }

    internal sealed class CalDavSyncStateStore
    {
        private const string FileName = "caldav-sync-state.xml";
        private readonly Dictionary<string, CalDavCalendarSyncState> _states =
            new Dictionary<string, CalDavCalendarSyncState>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, CalDavItemSyncState> _itemStates =
            new Dictionary<string, CalDavItemSyncState>(StringComparer.OrdinalIgnoreCase);

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

                foreach (XmlNode node in root.SelectNodes("Item"))
                {
                    XmlElement element = node as XmlElement;
                    if (element == null)
                    {
                        continue;
                    }
                    string calendarHref = (element.GetAttribute("calendarHref") ?? string.Empty).Trim();
                    string uid = (element.GetAttribute("uid") ?? string.Empty).Trim();
                    if (calendarHref.Length == 0 || uid.Length == 0)
                    {
                        continue;
                    }

                    DateTime parsedUtc;
                    string utcText = element.GetAttribute("lastSeenBothUtc") ?? string.Empty;
                    var itemState = new CalDavItemSyncState
                    {
                        CalendarHref = calendarHref,
                        Uid = uid,
                        Href = (element.GetAttribute("href") ?? string.Empty).Trim(),
                        ETag = (element.GetAttribute("etag") ?? string.Empty).Trim(),
                        LastSeenBothUtc = DateTime.TryParse(
                            utcText,
                            System.Globalization.CultureInfo.InvariantCulture,
                            System.Globalization.DateTimeStyles.RoundtripKind,
                            out parsedUtc)
                            ? (DateTime?)parsedUtc.ToUniversalTime()
                            : null
                    };
                    store._itemStates[BuildItemKey(calendarHref, uid)] = itemState;
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

        internal bool WasItemSeenOnBothSides(string calendarHref, string uid)
        {
            CalDavItemSyncState state = GetItem(calendarHref, uid);
            return state != null && state.LastSeenBothUtc.HasValue;
        }

        internal CalDavItemSyncState GetItem(string calendarHref, string uid)
        {
            string key = BuildItemKey(calendarHref, uid);
            CalDavItemSyncState state;
            return key.Length > 0 && _itemStates.TryGetValue(key, out state) ? state : null;
        }

        internal IList<CalDavItemSyncState> GetItems(string calendarHref)
        {
            string href = (calendarHref ?? string.Empty).Trim();
            var result = new List<CalDavItemSyncState>();
            foreach (CalDavItemSyncState state in _itemStates.Values)
            {
                if (state != null
                    && string.Equals(state.CalendarHref, href, StringComparison.OrdinalIgnoreCase))
                {
                    result.Add(state);
                }
            }
            return result;
        }

        internal void MarkItemSeenOnBothSides(
            string calendarHref,
            string uid,
            string href,
            string etag)
        {
            string calendar = (calendarHref ?? string.Empty).Trim();
            string cleanUid = (uid ?? string.Empty).Trim();
            if (calendar.Length == 0 || cleanUid.Length == 0)
            {
                return;
            }

            _itemStates[BuildItemKey(calendar, cleanUid)] = new CalDavItemSyncState
            {
                CalendarHref = calendar,
                Uid = cleanUid,
                Href = (href ?? string.Empty).Trim(),
                ETag = (etag ?? string.Empty).Trim(),
                LastSeenBothUtc = DateTime.UtcNow
            };
        }

        internal void RemoveItem(string calendarHref, string uid)
        {
            string key = BuildItemKey(calendarHref, uid);
            if (key.Length > 0)
            {
                _itemStates.Remove(key);
            }
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

                foreach (CalDavItemSyncState state in _itemStates.Values)
                {
                    if (state == null
                        || string.IsNullOrWhiteSpace(state.CalendarHref)
                        || string.IsNullOrWhiteSpace(state.Uid))
                    {
                        continue;
                    }
                    XmlElement item = document.CreateElement("Item");
                    item.SetAttribute("calendarHref", state.CalendarHref ?? string.Empty);
                    item.SetAttribute("uid", state.Uid ?? string.Empty);
                    item.SetAttribute("href", state.Href ?? string.Empty);
                    item.SetAttribute("etag", state.ETag ?? string.Empty);
                    if (state.LastSeenBothUtc.HasValue)
                    {
                        item.SetAttribute(
                            "lastSeenBothUtc",
                            state.LastSeenBothUtc.Value.ToUniversalTime().ToString("o", System.Globalization.CultureInfo.InvariantCulture));
                    }
                    root.AppendChild(item);
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

        private static string BuildItemKey(string calendarHref, string uid)
        {
            string calendar = (calendarHref ?? string.Empty).Trim();
            string cleanUid = (uid ?? string.Empty).Trim();
            if (calendar.Length == 0 || cleanUid.Length == 0)
            {
                return string.Empty;
            }
            return calendar + "|" + cleanUid;
        }

        private static string GetPath()
        {
            return Path.Combine(AppDataPaths.EnsureLocalRootDirectory(), FileName);
        }
    }
}
