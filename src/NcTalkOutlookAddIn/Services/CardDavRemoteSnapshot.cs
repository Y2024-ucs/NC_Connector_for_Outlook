// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Linq;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Services
{
    // Only a complete, validated addressbook-query may authorize local deletions.
    internal sealed class CardDavRemoteSnapshot
    {
        private static readonly XNamespace Dav = "DAV:";
        internal string AddressBookHref { get; private set; }
        internal Dictionary<string, string> Versions { get; private set; }

        internal static CardDavRemoteSnapshot Parse(string addressBookHref, XDocument document)
        {
            var book = new Uri(addressBookHref.TrimEnd('/') + "/", UriKind.Absolute);
            if (document.Root == null || document.Root.Name != Dav + "multistatus"
                || document.Descendants(Dav + "error").Any()
                || document.Root.Elements().Any(e => e.Name != Dav + "response"
                    && e.Name != Dav + "responsedescription"))
            {
                throw new InvalidOperationException("Incomplete CardDAV address book response.");
            }
            var result = new CardDavRemoteSnapshot
            {
                AddressBookHref = book.AbsoluteUri,
                Versions = new Dictionary<string, string>(StringComparer.Ordinal)
            };
            foreach (XElement response in document.Root.Elements(Dav + "response"))
            {
                string href = ((string)response.Element(Dav + "href") ?? string.Empty).Trim();
                Uri contact;
                if (href.Length == 0 || !Uri.TryCreate(book, href, out contact)
                    || !result.ContainsResource(contact.AbsoluteUri)
                    || response.Element(Dav + "status") != null
                    || response.Elements(Dav + "propstat").Any(p => !Successful((string)p.Element(Dav + "status"))))
                {
                    throw new InvalidOperationException("Invalid or failed CardDAV contact listing.");
                }
                XElement etag = response.Elements(Dav + "propstat")
                    .Select(p => p.Element(Dav + "prop"))
                    .Where(p => p != null).Select(p => p.Element(Dav + "getetag"))
                    .FirstOrDefault(p => p != null);
                if (etag == null || string.IsNullOrWhiteSpace(etag.Value)
                    || result.Versions.ContainsKey(contact.AbsoluteUri))
                {
                    throw new InvalidOperationException("Incomplete or duplicate CardDAV contact version.");
                }
                result.Versions.Add(contact.AbsoluteUri, etag.Value.Trim());
            }
            return result;
        }

        internal bool ContainsResource(string href)
        {
            Uri uri;
            return Uri.TryCreate(href, UriKind.Absolute, out uri)
                && string.IsNullOrEmpty(uri.Query) && string.IsNullOrEmpty(uri.Fragment)
                && !uri.AbsolutePath.EndsWith("/", StringComparison.Ordinal)
                && string.Equals(new Uri(uri, ".").AbsoluteUri, AddressBookHref, StringComparison.Ordinal);
        }

        internal bool IsMissing(string href)
        {
            return ContainsResource(href) && !Versions.ContainsKey(new Uri(href).AbsoluteUri);
        }

        private static bool Successful(string status)
        {
            string[] parts = (status ?? string.Empty).Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return parts.Length >= 2 && parts[1] == "200";
        }
    }

    internal static class CardDavContactFolderSync
    {
        private const string HrefPropertyName = "NC-CardDAV-HREF";
        private const string GroupCopyPropertyName = "NC-CardDAV-GROUP-COPY";
        private const string GroupCopyNamePropertyName = "NC-CardDAV-GROUP-COPY-NAME";
        private const string GeneratedSystemAddressBookMarker = "z-server-generated--system";
        private const string PublicStringsBase = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/";
        private const string ManagedFolderProperty = PublicStringsBase + "NC-CardDAV-GROUP-FOLDER";
        private const string ManagedFolderNameProperty = PublicStringsBase + "NC-CardDAV-GROUP-FOLDER-NAME";
        private static readonly XNamespace Dav = "DAV:";
        private static readonly XNamespace CardDav = "urn:ietf:params:xml:ns:carddav";

        private sealed class SourceContactReference
        {
            internal string EntryId { get; set; }
            internal string StoreId { get; set; }
        }

        internal static int Reconcile(
            TalkServiceConfiguration configuration,
            Outlook.Application outlookApplication,
            IEnumerable<CardDavAddressBook> addressBooks,
            CardDavSyncPreferences preferences)
        {
            if (configuration == null || outlookApplication == null || addressBooks == null || preferences == null)
            {
                return 0;
            }

            Dictionary<string, List<string>> categoriesByHref = LoadRemoteCategories(configuration, addressBooks, preferences);
            var activeGroups = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (List<string> categories in categoriesByHref.Values)
            {
                if (categories == null) continue;
                foreach (string category in categories)
                {
                    string name = (category ?? string.Empty).Trim();
                    if (!string.IsNullOrWhiteSpace(name)) activeGroups.Add(name);
                }
            }

            Outlook.NameSpace session = null;
            Outlook.MAPIFolder defaultContacts = null;
            try
            {
                session = outlookApplication.Session;
                defaultContacts = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                Dictionary<string, SourceContactReference> sourceContacts = LoadSourceContacts(defaultContacts);
                DeleteObsoleteManagedFolders(defaultContacts, activeGroups);

                int folderCount = 0;
                foreach (string groupName in activeGroups.OrderBy(name => name, StringComparer.CurrentCultureIgnoreCase))
                {
                    Outlook.MAPIFolder groupFolder = null;
                    try
                    {
                        groupFolder = EnsureManagedGroupFolder(defaultContacts, groupName);
                        if (groupFolder == null) continue;

                        ClearManagedCopies(groupFolder);
                        int copied = CopyGroupMembers(session, groupFolder, groupName, sourceContacts, categoriesByHref);
                        folderCount++;
                        DiagnosticsLogger.Log(
                            LogCategories.Core,
                            "CardDAV contact group folder synchronized (group=" + groupName
                            + ", contacts=" + copied + ").");
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(groupFolder, LogCategories.Core, "Failed to release CardDAV group folder.");
                    }
                }

                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "CardDAV contact group folder reconciliation completed (folders=" + folderCount + ").");
                return folderCount;
            }
            finally
            {
                ComInteropScope.TryRelease(defaultContacts, LogCategories.Core, "Failed to release default contacts folder after CardDAV group folder sync.");
                ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session after CardDAV group folder sync.");
            }
        }

        private static Dictionary<string, SourceContactReference> LoadSourceContacts(Outlook.MAPIFolder defaultContacts)
        {
            var result = new Dictionary<string, SourceContactReference>(StringComparer.OrdinalIgnoreCase);
            CollectSourceContacts(defaultContacts, result);

            Outlook.Folders folders = null;
            try
            {
                folders = defaultContacts.Folders;
                int count = folders != null ? folders.Count : 0;
                for (int index = 1; index <= count; index++)
                {
                    Outlook.MAPIFolder folder = null;
                    try
                    {
                        folder = folders[index];
                        if (folder != null && !IsManagedGroupFolder(folder))
                        {
                            CollectSourceContacts(folder, result);
                        }
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(folder, LogCategories.Core, "Failed to release source contacts subfolder during CardDAV group folder scan.");
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(folders, LogCategories.Core, "Failed to release source contact folders during CardDAV group folder scan.");
            }
            return result;
        }

        private static void CollectSourceContacts(
            Outlook.MAPIFolder folder,
            IDictionary<string, SourceContactReference> result)
        {
            Outlook.Items items = null;
            try
            {
                items = folder.Items;
                string storeId = folder.StoreID;
                int count = items != null ? items.Count : 0;
                for (int index = 1; index <= count; index++)
                {
                    object raw = null;
                    Outlook.ContactItem contact = null;
                    try
                    {
                        raw = items[index];
                        contact = raw as Outlook.ContactItem;
                        if (contact == null) continue;

                        string href = ReadContactProperty(contact, HrefPropertyName);
                        if (string.IsNullOrWhiteSpace(href)
                            || string.IsNullOrWhiteSpace(contact.EntryID)
                            || result.ContainsKey(href.Trim()))
                        {
                            continue;
                        }
                        result[href.Trim()] = new SourceContactReference
                        {
                            EntryId = contact.EntryID,
                            StoreId = storeId
                        };
                    }
                    finally
                    {
                        if (contact != null)
                        {
                            ComInteropScope.TryRelease(contact, LogCategories.Core, "Failed to release source contact during CardDAV group folder scan.");
                        }
                        else
                        {
                            ComInteropScope.TryRelease(raw, LogCategories.Core, "Failed to release non-contact during CardDAV group folder scan.");
                        }
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(items, LogCategories.Core, "Failed to release source items during CardDAV group folder scan.");
            }
        }

        private static int CopyGroupMembers(
            Outlook.NameSpace session,
            Outlook.MAPIFolder targetFolder,
            string groupName,
            IDictionary<string, SourceContactReference> sourceContacts,
            IDictionary<string, List<string>> categoriesByHref)
        {
            int copiedCount = 0;
            foreach (KeyValuePair<string, List<string>> pair in categoriesByHref)
            {
                if (pair.Value == null || !pair.Value.Any(category => string.Equals(
                    (category ?? string.Empty).Trim(), groupName, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                SourceContactReference sourceReference;
                if (!sourceContacts.TryGetValue(pair.Key, out sourceReference)
                    || sourceReference == null
                    || string.IsNullOrWhiteSpace(sourceReference.EntryId))
                {
                    DiagnosticsLogger.Log(
                        LogCategories.Core,
                        "CardDAV group folder source contact missing (group=" + groupName + ", href=" + pair.Key + ").");
                    continue;
                }

                Outlook.ContactItem source = null;
                Outlook.ContactItem copy = null;
                object moved = null;
                try
                {
                    source = session.GetItemFromID(sourceReference.EntryId, sourceReference.StoreId) as Outlook.ContactItem;
                    if (source == null) continue;
                    copy = source.Copy() as Outlook.ContactItem;
                    if (copy == null) continue;

                    WriteContactProperty(copy, GroupCopyPropertyName, "1");
                    WriteContactProperty(copy, GroupCopyNamePropertyName, groupName);
                    copy.Save();
                    moved = copy.Move(targetFolder);
                    copiedCount++;
                }
                finally
                {
                    ComInteropScope.TryRelease(moved, LogCategories.Core, "Failed to release moved CardDAV group folder contact.");
                    ComInteropScope.TryRelease(copy, LogCategories.Core, "Failed to release copied CardDAV group folder contact.");
                    ComInteropScope.TryRelease(source, LogCategories.Core, "Failed to release source CardDAV group folder contact.");
                }
            }
            return copiedCount;
        }

        private static void ClearManagedCopies(Outlook.MAPIFolder folder)
        {
            Outlook.Items items = null;
            try
            {
                items = folder.Items;
                for (int index = items.Count; index >= 1; index--)
                {
                    object raw = null;
                    Outlook.ContactItem contact = null;
                    try
                    {
                        raw = items[index];
                        contact = raw as Outlook.ContactItem;
                        if (contact != null
                            && string.Equals(ReadContactProperty(contact, GroupCopyPropertyName), "1", StringComparison.Ordinal))
                        {
                            contact.Delete();
                        }
                    }
                    finally
                    {
                        if (contact != null)
                        {
                            ComInteropScope.TryRelease(contact, LogCategories.Core, "Failed to release managed CardDAV group folder contact during cleanup.");
                        }
                        else
                        {
                            ComInteropScope.TryRelease(raw, LogCategories.Core, "Failed to release non-contact during CardDAV group folder cleanup.");
                        }
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(items, LogCategories.Core, "Failed to release CardDAV group folder items during cleanup.");
            }
        }

        private static Outlook.MAPIFolder EnsureManagedGroupFolder(Outlook.MAPIFolder parent, string groupName)
        {
            Outlook.Folders folders = null;
            Outlook.MAPIFolder exactUnmanagedMatch = null;
            try
            {
                folders = parent.Folders;
                int count = folders != null ? folders.Count : 0;
                for (int index = 1; index <= count; index++)
                {
                    Outlook.MAPIFolder folder = null;
                    try
                    {
                        folder = folders[index];
                        if (folder == null) continue;

                        string managedName = ReadFolderProperty(folder, ManagedFolderNameProperty);
                        if (IsManagedGroupFolder(folder)
                            && string.Equals(managedName, groupName, StringComparison.OrdinalIgnoreCase))
                        {
                            Outlook.MAPIFolder match = folder;
                            folder = null;
                            return match;
                        }

                        if (!IsManagedGroupFolder(folder)
                            && string.Equals(folder.Name, groupName, StringComparison.OrdinalIgnoreCase)
                            && exactUnmanagedMatch == null)
                        {
                            exactUnmanagedMatch = folder;
                            folder = null;
                        }
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(folder, LogCategories.Core, "Failed to release contact folder while locating CardDAV group folder.");
                    }
                }

                string folderName = SanitizeFolderName(groupName);
                if (exactUnmanagedMatch != null)
                {
                    folderName = SanitizeFolderName(groupName + " (Nextcloud)");
                    DiagnosticsLogger.Log(
                        LogCategories.Core,
                        "CardDAV group folder name collision; using alternate folder (group=" + groupName
                        + ", folder=" + folderName + ").");
                }

                Outlook.MAPIFolder created = folders.Add(folderName, Outlook.OlDefaultFolders.olFolderContacts);
                WriteFolderProperty(created, ManagedFolderProperty, "1");
                WriteFolderProperty(created, ManagedFolderNameProperty, groupName);
                return created;
            }
            finally
            {
                ComInteropScope.TryRelease(exactUnmanagedMatch, LogCategories.Core, "Failed to release unmanaged folder name collision reference.");
                ComInteropScope.TryRelease(folders, LogCategories.Core, "Failed to release folders collection after CardDAV group folder creation.");
            }
        }

        private static void DeleteObsoleteManagedFolders(Outlook.MAPIFolder parent, ISet<string> activeGroups)
        {
            Outlook.Folders folders = null;
            try
            {
                folders = parent.Folders;
                for (int index = folders.Count; index >= 1; index--)
                {
                    Outlook.MAPIFolder folder = null;
                    try
                    {
                        folder = folders[index];
                        if (folder == null || !IsManagedGroupFolder(folder)) continue;

                        string groupName = ReadFolderProperty(folder, ManagedFolderNameProperty);
                        if (string.IsNullOrWhiteSpace(groupName) || !activeGroups.Contains(groupName))
                        {
                            folder.Delete();
                        }
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(folder, LogCategories.Core, "Failed to release obsolete CardDAV group folder.");
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(folders, LogCategories.Core, "Failed to release folders collection after CardDAV group folder cleanup.");
            }
        }

        private static bool IsManagedGroupFolder(Outlook.MAPIFolder folder)
        {
            return string.Equals(ReadFolderProperty(folder, ManagedFolderProperty), "1", StringComparison.Ordinal);
        }

        private static string ReadFolderProperty(Outlook.MAPIFolder folder, string schemaName)
        {
            Outlook.PropertyAccessor accessor = null;
            try
            {
                accessor = folder != null ? folder.PropertyAccessor : null;
                if (accessor == null) return string.Empty;
                object value = accessor.GetProperty(schemaName);
                return value != null ? Convert.ToString(value) ?? string.Empty : string.Empty;
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                ComInteropScope.TryRelease(accessor, LogCategories.Core, "Failed to release CardDAV group folder property accessor.");
            }
        }

        private static void WriteFolderProperty(Outlook.MAPIFolder folder, string schemaName, string value)
        {
            Outlook.PropertyAccessor accessor = null;
            try
            {
                accessor = folder.PropertyAccessor;
                accessor.SetProperty(schemaName, value ?? string.Empty);
            }
            finally
            {
                ComInteropScope.TryRelease(accessor, LogCategories.Core, "Failed to release CardDAV group folder property accessor after write.");
            }
        }

        private static string ReadContactProperty(Outlook.ContactItem item, string name)
        {
            Outlook.UserProperties properties = null;
            Outlook.UserProperty property = null;
            try
            {
                properties = item != null ? item.UserProperties : null;
                property = properties != null ? properties.Find(name, true) : null;
                return property != null ? Convert.ToString(property.Value) ?? string.Empty : string.Empty;
            }
            finally
            {
                ComInteropScope.TryRelease(property, LogCategories.Core, "Failed to release CardDAV group folder contact property.");
                ComInteropScope.TryRelease(properties, LogCategories.Core, "Failed to release CardDAV group folder contact properties.");
            }
        }

        private static void WriteContactProperty(Outlook.ContactItem item, string name, string value)
        {
            Outlook.UserProperties properties = null;
            Outlook.UserProperty property = null;
            try
            {
                properties = item.UserProperties;
                property = properties.Find(name, true);
                if (property == null)
                {
                    property = properties.Add(name, Outlook.OlUserPropertyType.olText, true, Type.Missing);
                }
                property.Value = value ?? string.Empty;
            }
            finally
            {
                ComInteropScope.TryRelease(property, LogCategories.Core, "Failed to release CardDAV group folder contact property after write.");
                ComInteropScope.TryRelease(properties, LogCategories.Core, "Failed to release CardDAV group folder contact properties after write.");
            }
        }

        private static Dictionary<string, List<string>> LoadRemoteCategories(
            TalkServiceConfiguration configuration,
            IEnumerable<CardDavAddressBook> addressBooks,
            CardDavSyncPreferences preferences)
        {
            var result = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            foreach (CardDavAddressBook addressBook in addressBooks)
            {
                bool isSystem = IsSystemAddressBook(addressBook);
                if ((isSystem && !preferences.SyncCompanyDirectory)
                    || (!isSystem && !preferences.SyncPersonalContacts))
                {
                    continue;
                }

                XDocument document = SendCategoriesQuery(configuration, addressBook);
                foreach (XElement responseElement in document.Descendants(Dav + "response"))
                {
                    XElement prop = responseElement
                        .Elements(Dav + "propstat")
                        .Where(p => IsSuccessfulStatus((string)p.Element(Dav + "status")))
                        .Select(p => p.Element(Dav + "prop"))
                        .FirstOrDefault(p => p != null);
                    if (prop == null) continue;

                    string requestHref = ((string)responseElement.Element(Dav + "href") ?? string.Empty).Trim();
                    string href = ResolveUri(addressBook.Href, requestHref);
                    string vcard = (string)prop.Element(CardDav + "address-data");
                    if (string.IsNullOrWhiteSpace(href) || string.IsNullOrWhiteSpace(vcard)) continue;
                    result[href] = ParseCategories(vcard);
                }
            }
            return result;
        }

        private static XDocument SendCategoriesQuery(
            TalkServiceConfiguration configuration,
            CardDavAddressBook addressBook)
        {
            if (addressBook == null || string.IsNullOrWhiteSpace(addressBook.Href))
            {
                throw new InvalidOperationException("CardDAV address book is missing for group folder sync.");
            }

            string body = "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                + "<card:addressbook-query xmlns:d=\"DAV:\" xmlns:card=\"urn:ietf:params:xml:ns:carddav\">"
                + "<d:prop><card:address-data content-type=\"text/vcard\">"
                + "<card:prop name=\"CATEGORIES\"/>"
                + "</card:address-data></d:prop><card:filter/></card:addressbook-query>";

            var client = new NcHttpClient(configuration);
            var headers = new Dictionary<string, string>();
            headers["Depth"] = "1";
            NcHttpResponse response = client.Send(new NcHttpRequestOptions
            {
                Method = "REPORT",
                Url = addressBook.Href,
                Payload = body,
                Accept = "application/xml, text/xml",
                ContentType = "application/xml; charset=utf-8",
                IncludeOcsApiHeader = false,
                ParseJson = false,
                Headers = headers
            });

            if (!response.HasHttpResponse || (int)response.StatusCode != 207)
            {
                throw new InvalidOperationException("CardDAV group folder query failed for " + addressBook.Href + ".");
            }

            XDocument document = XDocument.Parse(response.ResponseText ?? string.Empty);
            if (document.Root == null || document.Root.Name != Dav + "multistatus"
                || document.Descendants(Dav + "error").Any()
                || document.Descendants(Dav + "status").Any(status => !IsSuccessfulStatus(status.Value)))
            {
                throw new InvalidOperationException("Incomplete CardDAV group folder response for " + addressBook.Href + ".");
            }
            return document;
        }

        private static List<string> ParseCategories(string rawVCard)
        {
            string normalized = (rawVCard ?? string.Empty)
                .Replace("\r\n", "\n")
                .Replace("\r", "\n");
            string[] sourceLines = normalized.Split('\n');
            var lines = new List<string>();
            foreach (string sourceLine in sourceLines)
            {
                if ((sourceLine.StartsWith(" ") || sourceLine.StartsWith("\t")) && lines.Count > 0)
                {
                    lines[lines.Count - 1] += sourceLine.Substring(1);
                }
                else
                {
                    lines.Add(sourceLine);
                }
            }

            var result = new List<string>();
            foreach (string line in lines)
            {
                int colon = line.IndexOf(':');
                if (colon <= 0) continue;
                string field = line.Substring(0, colon);
                string name = field.Split(';')[0];
                if (!string.Equals(name, "CATEGORIES", StringComparison.OrdinalIgnoreCase)) continue;

                foreach (string category in SplitCategoryValues(line.Substring(colon + 1)))
                {
                    string value = (category ?? string.Empty).Trim();
                    if (!string.IsNullOrWhiteSpace(value)
                        && !result.Any(existing => string.Equals(existing, value, StringComparison.OrdinalIgnoreCase)))
                    {
                        result.Add(value);
                    }
                }
            }
            return result;
        }

        private static IEnumerable<string> SplitCategoryValues(string value)
        {
            var result = new List<string>();
            var current = new StringBuilder();
            bool escaped = false;
            foreach (char c in value ?? string.Empty)
            {
                if (escaped)
                {
                    if (c == 'n' || c == 'N') current.Append(Environment.NewLine);
                    else current.Append(c);
                    escaped = false;
                }
                else if (c == '\\') escaped = true;
                else if (c == ',')
                {
                    result.Add(current.ToString());
                    current.Clear();
                }
                else current.Append(c);
            }
            if (escaped) current.Append('\\');
            result.Add(current.ToString());
            return result;
        }

        private static bool IsSystemAddressBook(CardDavAddressBook addressBook)
        {
            return addressBook != null
                && !string.IsNullOrWhiteSpace(addressBook.Href)
                && addressBook.Href.IndexOf(GeneratedSystemAddressBookMarker, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool IsSuccessfulStatus(string status)
        {
            return !string.IsNullOrWhiteSpace(status)
                && status.IndexOf(" 200 ", StringComparison.Ordinal) >= 0;
        }

        private static string ResolveUri(string baseUrl, string href)
        {
            Uri baseUri = new Uri(baseUrl, UriKind.Absolute);
            Uri absolute;
            return Uri.TryCreate(href, UriKind.Absolute, out absolute)
                ? absolute.AbsoluteUri
                : new Uri(baseUri, href ?? string.Empty).AbsoluteUri;
        }

        private static string SanitizeFolderName(string value)
        {
            string result = (value ?? string.Empty).Trim();
            foreach (char c in new[] { '\\', '/', ':', '*', '?', '"', '<', '>', '|' })
            {
                result = result.Replace(c, '-');
            }
            return string.IsNullOrWhiteSpace(result) ? "Nextcloud-Gruppe" : result;
        }
    }
}
