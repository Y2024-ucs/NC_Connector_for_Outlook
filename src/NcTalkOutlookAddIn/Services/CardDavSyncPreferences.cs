// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.IO;
using System.Xml;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class CardDavSyncPreferences
    {
        private const string FileName = "carddav-sync.xml";

        internal CardDavSyncPreferences()
        {
            Configured = false;
            Enabled = true;
            SyncCompanyDirectory = true;
            SyncPersonalContacts = true;
            UseDefaultContactsFolder = false;
        }

        internal bool Configured { get; set; }
        internal bool Enabled { get; set; }
        internal bool SyncCompanyDirectory { get; set; }
        internal bool SyncPersonalContacts { get; set; }
        internal bool UseDefaultContactsFolder { get; set; }

        internal static CardDavSyncPreferences Load()
        {
            var preferences = new CardDavSyncPreferences();
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
                if (root == null || !string.Equals(root.Name, "CardDavSync", StringComparison.OrdinalIgnoreCase))
                {
                    return preferences;
                }

                preferences.Configured = ReadBool(root, "Configured", false);
                preferences.Enabled = ReadBool(root, "Enabled", true);
                preferences.SyncCompanyDirectory = ReadBool(root, "SyncCompanyDirectory", true);
                preferences.SyncPersonalContacts = ReadBool(root, "SyncPersonalContacts", true);
                preferences.UseDefaultContactsFolder = ReadBool(root, "UseDefaultContactsFolder", false);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to load CardDAV sync preferences.", ex);
            }
            return preferences;
        }

        internal void Save()
        {
            string path = GetPath();
            try
            {
                var document = new XmlDocument();
                XmlElement root = document.CreateElement("CardDavSync");
                document.AppendChild(root);
                AppendBool(document, root, "Configured", Configured);
                AppendBool(document, root, "Enabled", Enabled);
                AppendBool(document, root, "SyncCompanyDirectory", SyncCompanyDirectory);
                AppendBool(document, root, "SyncPersonalContacts", SyncPersonalContacts);
                AppendBool(document, root, "UseDefaultContactsFolder", UseDefaultContactsFolder);

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
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to save CardDAV sync preferences.", ex);
                throw;
            }
        }

        private static string GetPath()
        {
            return Path.Combine(AppDataPaths.EnsureLocalRootDirectory(), FileName);
        }

        private static bool ReadBool(XmlElement root, string name, bool fallback)
        {
            XmlElement element = root[name];
            bool value;
            return element != null && bool.TryParse(element.InnerText, out value) ? value : fallback;
        }

        private static void AppendBool(XmlDocument document, XmlElement root, string name, bool value)
        {
            XmlElement element = document.CreateElement(name);
            element.InnerText = value.ToString(System.Globalization.CultureInfo.InvariantCulture);
            root.AppendChild(element);
        }
    }
}
