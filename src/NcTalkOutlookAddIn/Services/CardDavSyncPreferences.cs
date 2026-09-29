// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class CardDavSyncPreferences
    {
        private const string FileName = "carddav-sync.xml";
        private const string SystemAddressBookMarker = "z-server-generated--system";
        // Selection value for "all personal address books".
        internal const string AllPersonalAddressBooks = "*";
        // IBP default: the book synced from the Fritz!Box, shown as "IBP-Kontakte".
        internal const string DefaultAddressBookSlug = "fritzbox-kontakte";
        internal const string DefaultAddressBookDisplayName = "IBP-Kontakte";

        internal CardDavSyncPreferences()
        {
            Configured = false;
            Enabled = true;
            SyncPersonalContacts = true;
            PersonalAddressBookHref = string.Empty;
            PersonalAddressBookName = string.Empty;
            UseDefaultContactsFolder = true;
            SyncCustomers = true;
        }

        internal bool Configured { get; set; }
        internal bool Enabled { get; set; }
        internal bool SyncPersonalContacts { get; set; }
        // Empty: the IBP default book (all books when it does not exist); "*": all personal books;
        // otherwise only the book with this href.
        internal string PersonalAddressBookHref { get; set; }
        // Display name of the selected book, shown while the server list is not loaded.
        internal string PersonalAddressBookName { get; set; }
        internal bool UseDefaultContactsFolder { get; set; }
        // Two-way sync of the Outlook folder "Kunden" with the Nextcloud address book "IBP-Kunden".
        internal bool SyncCustomers { get; set; }

        /// <summary>
        /// Whether the read-only sync imports this address book, given all books of the account. The company
        /// directory (system address book) and the two-way customer book are never part of it.
        /// </summary>
        internal bool IncludesAddressBook(CardDavAddressBook addressBook, IEnumerable<CardDavAddressBook> allAddressBooks)
        {
            if (!IsPersonalAddressBook(addressBook) || !SyncPersonalContacts)
            {
                return false;
            }
            string selection = (PersonalAddressBookHref ?? string.Empty).Trim();
            if (string.Equals(selection, AllPersonalAddressBooks, StringComparison.Ordinal))
            {
                return true;
            }
            if (selection.Length > 0)
            {
                return SameCollection(addressBook.Href, selection);
            }
            bool defaultExists = (allAddressBooks ?? Enumerable.Empty<CardDavAddressBook>())
                .Any(book => IsPersonalAddressBook(book) && IsDefaultAddressBook(book));
            return !defaultExists || IsDefaultAddressBook(addressBook);
        }

        internal static bool IsPersonalAddressBook(CardDavAddressBook addressBook)
        {
            return addressBook != null
                && !string.IsNullOrWhiteSpace(addressBook.Href)
                && !IsSystemAddressBook(addressBook)
                && !CardDavCustomerSync.IsCustomerAddressBook(addressBook);
        }

        internal static bool IsDefaultAddressBook(CardDavAddressBook addressBook)
        {
            if (addressBook == null) return false;
            Uri uri;
            if (Uri.TryCreate(addressBook.Href ?? string.Empty, UriKind.Absolute, out uri)
                && string.Equals(uri.AbsolutePath.TrimEnd('/').Split('/').Last(), DefaultAddressBookSlug, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            return string.Equals((addressBook.DisplayName ?? string.Empty).Trim(), DefaultAddressBookDisplayName, StringComparison.OrdinalIgnoreCase);
        }

        internal static bool IsSystemAddressBook(CardDavAddressBook addressBook)
        {
            return addressBook != null
                && !string.IsNullOrWhiteSpace(addressBook.Href)
                && addressBook.Href.IndexOf(SystemAddressBookMarker, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        internal static bool SameCollection(string left, string right)
        {
            return string.Equals(
                (left ?? string.Empty).Trim().TrimEnd('/'),
                (right ?? string.Empty).Trim().TrimEnd('/'),
                StringComparison.OrdinalIgnoreCase);
        }

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
                preferences.SyncPersonalContacts = ReadBool(root, "SyncPersonalContacts", true);
                preferences.PersonalAddressBookHref = ReadString(root, "PersonalAddressBookHref");
                preferences.PersonalAddressBookName = ReadString(root, "PersonalAddressBookName");
                preferences.UseDefaultContactsFolder = ReadBool(root, "UseDefaultContactsFolder", true);
                preferences.SyncCustomers = ReadBool(root, "SyncCustomers", true);
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
                AppendBool(document, root, "SyncPersonalContacts", SyncPersonalContacts);
                AppendString(document, root, "PersonalAddressBookHref", PersonalAddressBookHref);
                AppendString(document, root, "PersonalAddressBookName", PersonalAddressBookName);
                AppendBool(document, root, "UseDefaultContactsFolder", UseDefaultContactsFolder);
                AppendBool(document, root, "SyncCustomers", SyncCustomers);

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

        private static string ReadString(XmlElement root, string name)
        {
            XmlElement element = root[name];
            return element != null ? (element.InnerText ?? string.Empty).Trim() : string.Empty;
        }

        private static void AppendBool(XmlDocument document, XmlElement root, string name, bool value)
        {
            XmlElement element = document.CreateElement(name);
            element.InnerText = value.ToString(System.Globalization.CultureInfo.InvariantCulture);
            root.AppendChild(element);
        }

        private static void AppendString(XmlDocument document, XmlElement root, string name, string value)
        {
            XmlElement element = document.CreateElement(name);
            element.InnerText = value ?? string.Empty;
            root.AppendChild(element);
        }
    }
}
