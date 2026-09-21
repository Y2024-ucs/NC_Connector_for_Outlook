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
    internal sealed class CardDavContactRecord
    {
        internal string Href { get; set; }
        internal string ETag { get; set; }
        internal string Uid { get; set; }
        internal string FullName { get; set; }
        internal string FirstName { get; set; }
        internal string LastName { get; set; }
        internal string Company { get; set; }
        internal string JobTitle { get; set; }
        internal string Email1 { get; set; }
        internal string Email2 { get; set; }
        internal string BusinessPhone { get; set; }
        internal string BusinessPhone2 { get; set; }
        internal string MobilePhone { get; set; }
        internal string HomePhone { get; set; }
        internal string HomePhone2 { get; set; }
        internal string BusinessFax { get; set; }
        internal string HomeFax { get; set; }
        internal string OtherPhone { get; set; }
        internal string PagerPhone { get; set; }
        internal string CompanyMainPhone { get; set; }
        internal string CarPhone { get; set; }
        internal string BusinessAddressStreet { get; set; }
        internal string BusinessAddressCity { get; set; }
        internal string BusinessAddressState { get; set; }
        internal string BusinessAddressPostalCode { get; set; }
        internal string BusinessAddressCountry { get; set; }
        internal string Notes { get; set; }
    }

    internal sealed class CardDavReadOnlySync
    {
        private sealed class RemoteContactVersion
        {
            internal string RequestHref { get; set; }
            internal string Href { get; set; }
            internal string ETag { get; set; }
        }

        private const string UidPropertyName = "NC-CardDAV-UID";
        private const string HrefPropertyName = "NC-CardDAV-HREF";
        private const string ETagPropertyName = "NC-CardDAV-ETAG";
        private const string ImportSchemaPropertyName = "NC-CardDAV-SCHEMA";
        private const string CurrentImportSchema = "2";
        private const string CompanyFolderSuffix = " - Firmenverzeichnis";
        private const string PersonalFolderSuffix = " - Persönliche Kontakte";
        private const string GeneratedSystemAddressBookMarker = "z-server-generated--system";

        private static readonly XNamespace Dav = "DAV:";
        private static readonly XNamespace CardDav = "urn:ietf:params:xml:ns:carddav";
        private static readonly object KnownEtagsSyncRoot = new object();
        private static readonly Dictionary<string, string> KnownEtags =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        private readonly TalkServiceConfiguration _configuration;
        private readonly List<CardDavRemoteSnapshot> _completedSnapshots = new List<CardDavRemoteSnapshot>();
        internal int DeletedCount { get; private set; }

        internal CardDavReadOnlySync(TalkServiceConfiguration configuration)
        {
            if (configuration == null)
            {
                throw new ArgumentNullException("configuration");
            }
            _configuration = configuration;
        }

        internal IList<CardDavContactRecord> DownloadContacts(CardDavAddressBook addressBook)
        {
            Dictionary<string, string> snapshot;
            lock (KnownEtagsSyncRoot)
            {
                snapshot = new Dictionary<string, string>(KnownEtags, StringComparer.OrdinalIgnoreCase);
            }
            return DownloadContacts(addressBook, snapshot);
        }

        internal IList<CardDavContactRecord> DownloadContacts(
            CardDavAddressBook addressBook,
            IDictionary<string, string> knownEtags)
        {
            if (addressBook == null || string.IsNullOrWhiteSpace(addressBook.Href))
            {
                throw new ArgumentException("A CardDAV address book is required.", "addressBook");
            }

            CardDavRemoteSnapshot snapshot = DownloadContactVersions(addressBook);
            IList<RemoteContactVersion> remoteVersions = snapshot.Versions.Select(pair => new RemoteContactVersion
            {
                RequestHref = pair.Key,
                Href = pair.Key,
                ETag = pair.Value
            }).ToList();
            var changedVersions = new List<RemoteContactVersion>();
            int unchangedCount = 0;

            foreach (RemoteContactVersion version in remoteVersions)
            {
                string knownEtag;
                if (!string.IsNullOrWhiteSpace(version.Href)
                    && !string.IsNullOrWhiteSpace(version.ETag)
                    && knownEtags != null
                    && knownEtags.TryGetValue(version.Href, out knownEtag)
                    && string.Equals(
                        NormalizeEtag(knownEtag),
                        NormalizeEtag(version.ETag),
                        StringComparison.Ordinal))
                {
                    RememberKnownEtag(version.Href, version.ETag);
                    unchangedCount++;
                    continue;
                }
                changedVersions.Add(version);
            }

            if (changedVersions.Count == 0)
            {
                _completedSnapshots.Add(snapshot);
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "CardDAV change scan completed (addressBook="
                    + (addressBook.DisplayName ?? string.Empty)
                    + ", remote=" + remoteVersions.Count
                    + ", changed=0, unchanged=" + unchangedCount
                    + ", payloadFetched=0). ");
                return new List<CardDavContactRecord>();
            }

            IList<CardDavContactRecord> contacts = DownloadContactPayloads(addressBook, changedVersions);
            var expected = new HashSet<string>(changedVersions.Select(v => v.Href), StringComparer.Ordinal);
            if (contacts.Count != expected.Count || contacts.Any(c => !expected.Remove(c.Href)) || expected.Count != 0)
            {
                throw new InvalidOperationException("Incomplete CardDAV contact payload response; deletion skipped.");
            }
            _completedSnapshots.Add(snapshot);
            DiagnosticsLogger.Log(
                LogCategories.Core,
                "CardDAV change scan completed (addressBook="
                + (addressBook.DisplayName ?? string.Empty)
                + ", remote=" + remoteVersions.Count
                + ", changed=" + changedVersions.Count
                + ", unchanged=" + unchangedCount
                + ", payloadFetched=" + contacts.Count
                + ").");
            return contacts;
        }

        private CardDavRemoteSnapshot DownloadContactVersions(CardDavAddressBook addressBook)
        {
            string body = "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                + "<card:addressbook-query xmlns:d=\"DAV:\" xmlns:card=\"urn:ietf:params:xml:ns:carddav\">"
                + "<d:prop><d:getetag/></d:prop><card:filter/></card:addressbook-query>";
            return CardDavRemoteSnapshot.Parse(addressBook.Href,
                SendDavRequest(addressBook.Href, "REPORT", "1", body));
        }

        private IList<CardDavContactRecord> DownloadContactPayloads(
            CardDavAddressBook addressBook,
            IList<RemoteContactVersion> changedVersions)
        {
            if (changedVersions == null || changedVersions.Count == 0)
            {
                return new List<CardDavContactRecord>();
            }

            var body = new StringBuilder();
            body.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            body.Append("<card:addressbook-multiget xmlns:d=\"DAV:\" xmlns:card=\"urn:ietf:params:xml:ns:carddav\">");
            body.Append("<d:prop><d:getetag/><card:address-data content-type=\"text/vcard\"/></d:prop>");
            foreach (RemoteContactVersion version in changedVersions)
            {
                body.Append("<d:href>");
                body.Append(EscapeXml(version.RequestHref));
                body.Append("</d:href>");
            }
            body.Append("</card:addressbook-multiget>");

            XDocument document = SendDavRequest(addressBook.Href, "REPORT", "1", body.ToString());
            var contacts = new List<CardDavContactRecord>();
            foreach (XElement responseElement in document.Descendants(Dav + "response"))
            {
                XElement prop = responseElement
                    .Elements(Dav + "propstat")
                    .Where(p => IsSuccessfulStatus((string)p.Element(Dav + "status")))
                    .Select(p => p.Element(Dav + "prop"))
                    .FirstOrDefault(p => p != null);
                if (prop == null)
                {
                    continue;
                }

                string requestHref = ((string)responseElement.Element(Dav + "href") ?? string.Empty).Trim();
                string href = ResolveUri(addressBook.Href, requestHref);
                string etag = ((string)prop.Element(Dav + "getetag") ?? string.Empty).Trim();
                string vcard = (string)prop.Element(CardDav + "address-data");
                if (string.IsNullOrWhiteSpace(vcard))
                {
                    continue;
                }

                vcard = CardDavVCardSupport.Normalize(vcard);
                CardDavContactRecord contact = ParseVCard(vcard);
                contact.Href = href;
                contact.ETag = etag;
                CardDavVCardSupport.CachePhoto(contact.Href, vcard, _configuration);
                contacts.Add(contact);
            }
            return contacts;
        }

        private static string EscapeXml(string value)
        {
            return (value ?? string.Empty)
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;")
                .Replace("'", "&apos;");
        }

        internal static Dictionary<string, string> LoadKnownEtagsFromOutlook(
            Outlook.Application outlookApplication)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (outlookApplication == null)
            {
                return result;
            }

            Outlook.NameSpace session = null;
            Outlook.MAPIFolder defaultContacts = null;
            Outlook.Folders folders = null;
            try
            {
                session = outlookApplication.Session;
                defaultContacts = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                CollectKnownEtagsFromFolder(defaultContacts, result);

                folders = defaultContacts.Folders;
                int count = folders.Count;
                for (int index = 1; index <= count; index++)
                {
                    Outlook.MAPIFolder folder = null;
                    try
                    {
                        folder = folders[index];
                        if (folder != null)
                        {
                            CollectKnownEtagsFromFolder(folder, result);
                        }
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(folder, LogCategories.Core, "Failed to release CardDAV contacts subfolder during ETag scan.");
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(folders, LogCategories.Core, "Failed to release CardDAV contacts subfolders during ETag scan.");
                ComInteropScope.TryRelease(defaultContacts, LogCategories.Core, "Failed to release default contacts folder during CardDAV ETag scan.");
                ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session after CardDAV ETag scan.");
            }

            lock (KnownEtagsSyncRoot)
            {
                KnownEtags.Clear();
                foreach (KeyValuePair<string, string> pair in result)
                {
                    KnownEtags[pair.Key] = pair.Value;
                }
            }

            DiagnosticsLogger.Log(
                LogCategories.Core,
                "CardDAV local ETag cache initialized (contacts=" + result.Count + ").");
            return result;
        }

        private static void CollectKnownEtagsFromFolder(
            Outlook.MAPIFolder folder,
            IDictionary<string, string> result)
        {
            if (folder == null || result == null)
            {
                return;
            }

            Outlook.Items items = null;
            try
            {
                items = folder.Items;
                int count = items.Count;
                for (int index = 1; index <= count; index++)
                {
                    object raw = null;
                    Outlook.ContactItem contact = null;
                    try
                    {
                        raw = items[index];
                        contact = raw as Outlook.ContactItem;
                        if (contact == null)
                        {
                            continue;
                        }

                        string href = ReadUserProperty(contact, HrefPropertyName);
                        string etag = ReadUserProperty(contact, ETagPropertyName);
                        string importSchema = ReadUserProperty(contact, ImportSchemaPropertyName);
                        if (!string.IsNullOrWhiteSpace(href)
                            && !string.IsNullOrWhiteSpace(etag)
                            && string.Equals(
                                importSchema,
                                CurrentImportSchema,
                                StringComparison.Ordinal))
                        {
                            result[href.Trim()] = etag.Trim();
                        }
                    }
                    finally
                    {
                        if (contact != null)
                        {
                            ComInteropScope.TryRelease(contact, LogCategories.Core, "Failed to release CardDAV contact during ETag scan.");
                        }
                        else
                        {
                            ComInteropScope.TryRelease(raw, LogCategories.Core, "Failed to release non-contact item during CardDAV ETag scan.");
                        }
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(items, LogCategories.Core, "Failed to release CardDAV contact items during ETag scan.");
            }
        }

        internal static void RememberKnownEtag(string href, string etag)
        {
            if (string.IsNullOrWhiteSpace(href) || string.IsNullOrWhiteSpace(etag))
            {
                return;
            }
            lock (KnownEtagsSyncRoot)
            {
                KnownEtags[href.Trim()] = etag.Trim();
            }
        }

        private static string NormalizeEtag(string value)
        {
            return (value ?? string.Empty).Trim();
        }

        internal int ImportIntoOutlook(
            Outlook.Application outlookApplication,
            IEnumerable<CardDavContactRecord> contacts,
            bool? useDefaultContactsFolder = null)
        {
            if (outlookApplication == null)
            {
                throw new ArgumentNullException("outlookApplication");
            }

            CardDavBackgroundSyncManager.EnsureStarted(_configuration, outlookApplication);

            CardDavSyncPreferences preferences = CardDavSyncPreferences.Load();
            if (useDefaultContactsFolder ?? preferences.UseDefaultContactsFolder)
            {
                int imported = CardDavDefaultContactsImporter.Import(outlookApplication, contacts);
                ReconcileDeletedContacts(outlookApplication);
                return imported;
            }

            List<CardDavContactRecord> sourceContacts = (contacts ?? Enumerable.Empty<CardDavContactRecord>())
                .Where(c => c != null)
                .ToList();
            List<CardDavContactRecord> companyContacts = sourceContacts
                .Where(IsSystemDirectoryContact)
                .ToList();
            List<CardDavContactRecord> personalContacts = sourceContacts
                .Where(c => !IsSystemDirectoryContact(c))
                .ToList();

            Outlook.NameSpace session = null;
            Outlook.MAPIFolder defaultContacts = null;
            Outlook.MAPIFolder companyFolder = null;
            Outlook.MAPIFolder personalFolder = null;
            try
            {
                session = outlookApplication.Session;
                defaultContacts = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                string instanceName = ResolveInstanceName();
                int imported = 0;

                if (companyContacts.Count > 0)
                {
                    companyFolder = EnsureSubFolder(
                        defaultContacts,
                        instanceName + CompanyFolderSuffix);
                    imported += ImportIntoFolder(session, companyFolder, companyContacts);
                }

                if (personalContacts.Count > 0)
                {
                    personalFolder = EnsureSubFolder(
                        defaultContacts,
                        instanceName + PersonalFolderSuffix);
                    imported += ImportIntoFolder(session, personalFolder, personalContacts);
                }

                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "CardDAV read-only sync imported/updated "
                    + imported
                    + " changed contacts (instance="
                    + instanceName
                    + ", company="
                    + companyContacts.Count
                    + ", personal="
                    + personalContacts.Count
                    + ").");
                ReconcileDeletedContacts(outlookApplication);
                return imported;
            }
            finally
            {
                ComInteropScope.TryRelease(personalFolder, LogCategories.Core, "Failed to release CardDAV personal folder.");
                ComInteropScope.TryRelease(companyFolder, LogCategories.Core, "Failed to release CardDAV company folder.");
                ComInteropScope.TryRelease(defaultContacts, LogCategories.Core, "Failed to release Outlook contacts folder.");
                ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session after CardDAV sync.");
            }
        }

        internal void ReconcileDeletedContacts(Outlook.Application application)
        {
            DeletedCount = 0;
            if (_completedSnapshots.Count == 0) return;
            Outlook.NameSpace session = null;
            Outlook.MAPIFolder root = null;
            Outlook.Folders folders = null;
            try
            {
                session = application.Session;
                root = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                DeletedCount += ReconcileFolder(root);
                folders = root.Folders;
                for (int index = 1; index <= folders.Count; index++)
                {
                    Outlook.MAPIFolder folder = null;
                    try
                    {
                        folder = folders[index];
                        if (folder.DefaultItemType == Outlook.OlItemType.olContactItem)
                            DeletedCount += ReconcileFolder(folder);
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(folder, LogCategories.Core, "Failed to release reconciled contacts folder.");
                    }
                }
                DiagnosticsLogger.Log(LogCategories.Core, "CardDAV reconciliation completed (deleted=" + DeletedCount + ").");
            }
            finally
            {
                ComInteropScope.TryRelease(folders, LogCategories.Core, "Failed to release reconciliation folders.");
                ComInteropScope.TryRelease(root, LogCategories.Core, "Failed to release reconciliation root.");
                ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release reconciliation session.");
            }
        }

        private int ReconcileFolder(Outlook.MAPIFolder folder)
        {
            Outlook.Items items = null;
            int deleted = 0;
            try
            {
                items = folder.Items;
                // Outlook collections are one-based; reverse traversal survives Delete().
                for (int index = items.Count; index >= 1; index--)
                {
                    object raw = null;
                    try
                    {
                        raw = items[index];
                        var contact = raw as Outlook.ContactItem;
                        if (contact == null) continue;
                        string href = ReadUserProperty(contact, HrefPropertyName).Trim();
                        if (!_completedSnapshots.Any(snapshot => snapshot.IsMissing(href))) continue;
                        contact.Delete();
                        deleted++;
                        lock (KnownEtagsSyncRoot) { KnownEtags.Remove(href); }
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(raw, LogCategories.Core, "Failed to release reconciled contact.");
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(items, LogCategories.Core, "Failed to release reconciliation items.");
            }
            return deleted;
        }

        private static int ImportIntoFolder(
            Outlook.NameSpace session,
            Outlook.MAPIFolder targetFolder,
            IEnumerable<CardDavContactRecord> contacts)
        {
            string storeId = targetFolder.StoreID;
            Dictionary<string, string> existing = LoadExistingContactEntryIds(targetFolder);
            int imported = 0;

            foreach (CardDavContactRecord source in contacts)
            {
                string key = BuildContactKey(source);
                Outlook.ContactItem target = null;
                Outlook.Items folderItems = null;
                try
                {
                    string entryId;
                    if (!string.IsNullOrWhiteSpace(key)
                        && existing.TryGetValue(key, out entryId)
                        && !string.IsNullOrWhiteSpace(entryId))
                    {
                        target = session.GetItemFromID(entryId, storeId) as Outlook.ContactItem;
                    }

                    if (target == null)
                    {
                        folderItems = targetFolder.Items;
                        target = folderItems.Add(Outlook.OlItemType.olContactItem)
                            as Outlook.ContactItem;
                    }
                    if (target == null)
                    {
                        continue;
                    }

                    string currentEtag = ReadUserProperty(target, ETagPropertyName);
                    string currentImportSchema = ReadUserProperty(target, ImportSchemaPropertyName);
                    if (!string.IsNullOrWhiteSpace(source.ETag)
                        && string.Equals(
                            NormalizeEtag(currentEtag),
                            NormalizeEtag(source.ETag),
                            StringComparison.Ordinal)
                        && string.Equals(
                            currentImportSchema,
                            CurrentImportSchema,
                            StringComparison.Ordinal))
                    {
                        continue;
                    }

                    ApplyContact(target, source);
                    target.Save();
                    RememberKnownEtag(source.Href, source.ETag);
                    imported++;
                }
                finally
                {
                    ComInteropScope.TryRelease(target, LogCategories.Core, "Failed to release CardDAV ContactItem.");
                    ComInteropScope.TryRelease(folderItems, LogCategories.Core, "Failed to release CardDAV Items collection.");
                }
            }
            return imported;
        }

        private static Dictionary<string, string> LoadExistingContactEntryIds(Outlook.MAPIFolder folder)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            Outlook.Items items = null;
            try
            {
                items = folder.Items;
                int count = items.Count;
                for (int index = 1; index <= count; index++)
                {
                    object raw = null;
                    Outlook.ContactItem contact = null;
                    try
                    {
                        raw = items[index];
                        contact = raw as Outlook.ContactItem;
                        if (contact == null)
                        {
                            continue;
                        }

                        string key = BuildContactKey(
                            ReadUserProperty(contact, UidPropertyName),
                            ReadUserProperty(contact, HrefPropertyName));
                        if (!string.IsNullOrWhiteSpace(key)
                            && !result.ContainsKey(key)
                            && !string.IsNullOrWhiteSpace(contact.EntryID))
                        {
                            result[key] = contact.EntryID;
                        }
                    }
                    finally
                    {
                        if (contact != null)
                        {
                            ComInteropScope.TryRelease(contact, LogCategories.Core, "Failed to release existing CardDAV ContactItem.");
                        }
                        else
                        {
                            ComInteropScope.TryRelease(raw, LogCategories.Core, "Failed to release non-contact CardDAV folder item.");
                        }
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(items, LogCategories.Core, "Failed to release CardDAV contact items collection.");
            }
            return result;
        }

        private static void ApplyContact(Outlook.ContactItem target, CardDavContactRecord source)
        {
            target.FullName = source.FullName ?? string.Empty;
            target.FirstName = source.FirstName ?? string.Empty;
            target.LastName = source.LastName ?? string.Empty;
            target.CompanyName = source.Company ?? string.Empty;
            target.JobTitle = source.JobTitle ?? string.Empty;
            target.Email1Address = source.Email1 ?? string.Empty;
            target.Email2Address = source.Email2 ?? string.Empty;
            target.BusinessTelephoneNumber = source.BusinessPhone ?? string.Empty;
            target.Business2TelephoneNumber = source.BusinessPhone2 ?? string.Empty;
            target.MobileTelephoneNumber = source.MobilePhone ?? string.Empty;
            target.HomeTelephoneNumber = source.HomePhone ?? string.Empty;
            target.Home2TelephoneNumber = source.HomePhone2 ?? string.Empty;
            target.BusinessFaxNumber = source.BusinessFax ?? string.Empty;
            target.HomeFaxNumber = source.HomeFax ?? string.Empty;
            target.OtherTelephoneNumber = source.OtherPhone ?? string.Empty;
            target.PagerNumber = source.PagerPhone ?? string.Empty;
            target.CompanyMainTelephoneNumber = source.CompanyMainPhone ?? string.Empty;
            target.CarTelephoneNumber = source.CarPhone ?? string.Empty;
            target.BusinessAddressStreet = source.BusinessAddressStreet ?? string.Empty;
            target.BusinessAddressCity = source.BusinessAddressCity ?? string.Empty;
            target.BusinessAddressState = source.BusinessAddressState ?? string.Empty;
            target.BusinessAddressPostalCode = source.BusinessAddressPostalCode ?? string.Empty;
            target.BusinessAddressCountry = source.BusinessAddressCountry ?? string.Empty;
            target.Body = source.Notes ?? string.Empty;

            WriteUserProperty(target, UidPropertyName, source.Uid);
            WriteUserProperty(target, HrefPropertyName, source.Href);
            WriteUserProperty(target, ETagPropertyName, source.ETag);
            WriteUserProperty(target, ImportSchemaPropertyName, CurrentImportSchema);
            CardDavVCardSupport.ApplyPhoto(target, source.Href);
        }

        private static bool IsSystemDirectoryContact(CardDavContactRecord contact)
        {
            return contact != null
                && !string.IsNullOrWhiteSpace(contact.Href)
                && contact.Href.IndexOf(
                    GeneratedSystemAddressBookMarker,
                    StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string BuildContactKey(CardDavContactRecord contact)
        {
            return contact == null
                ? string.Empty
                : BuildContactKey(contact.Uid, contact.Href);
        }

        private static string BuildContactKey(string uid, string href)
        {
            if (!string.IsNullOrWhiteSpace(uid))
            {
                return "uid:" + uid.Trim();
            }
            if (!string.IsNullOrWhiteSpace(href))
            {
                return "href:" + href.Trim();
            }
            return string.Empty;
        }

        private static string ReadUserProperty(Outlook.ContactItem item, string name)
        {
            Outlook.UserProperties properties = null;
            Outlook.UserProperty property = null;
            try
            {
                properties = item.UserProperties;
                property = properties != null ? properties.Find(name, true) : null;
                return property != null
                    ? Convert.ToString(property.Value) ?? string.Empty
                    : string.Empty;
            }
            finally
            {
                ComInteropScope.TryRelease(property, LogCategories.Core, "Failed to release CardDAV contact user property.");
                ComInteropScope.TryRelease(properties, LogCategories.Core, "Failed to release CardDAV contact user properties.");
            }
        }

        private static void WriteUserProperty(Outlook.ContactItem item, string name, string value)
        {
            Outlook.UserProperties properties = null;
            Outlook.UserProperty property = null;
            try
            {
                properties = item.UserProperties;
                if (properties == null)
                {
                    return;
                }
                property = properties.Find(name, true);
                if (property == null)
                {
                    property = properties.Add(
                        name,
                        Outlook.OlUserPropertyType.olText,
                        true,
                        Type.Missing);
                }
                property.Value = value ?? string.Empty;
            }
            finally
            {
                ComInteropScope.TryRelease(property, LogCategories.Core, "Failed to release CardDAV contact user property.");
                ComInteropScope.TryRelease(properties, LogCategories.Core, "Failed to release CardDAV contact user properties.");
            }
        }

        private static Outlook.MAPIFolder EnsureSubFolder(Outlook.MAPIFolder parent, string name)
        {
            Outlook.Folders folders = null;
            try
            {
                folders = parent.Folders;
                int count = folders.Count;
                for (int index = 1; index <= count; index++)
                {
                    Outlook.MAPIFolder folder = null;
                    try
                    {
                        folder = folders[index];
                        if (folder != null
                            && string.Equals(folder.Name, name, StringComparison.OrdinalIgnoreCase))
                        {
                            Outlook.MAPIFolder match = folder;
                            folder = null;
                            return match;
                        }
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(folder, LogCategories.Core, "Failed to release Outlook contact subfolder.");
                    }
                }
                return folders.Add(name, Outlook.OlDefaultFolders.olFolderContacts);
            }
            finally
            {
                ComInteropScope.TryRelease(folders, LogCategories.Core, "Failed to release Outlook contact folders collection.");
            }
        }

        private string ResolveInstanceName()
        {
            try
            {
                NextcloudCapabilitiesSnapshot snapshot = new NextcloudCapabilitiesService(_configuration)
                    .GetSnapshot(false, false);
                IDictionary<string, object> theming = NcJson.GetDictionary(
                    snapshot != null ? snapshot.Capabilities : null,
                    "theming");
                string themedName = NcJson.GetTrimmedString(theming, "name");
                if (!string.IsNullOrWhiteSpace(themedName)
                    && !string.Equals(themedName, "Nextcloud", StringComparison.OrdinalIgnoreCase))
                {
                    return SanitizeFolderName(themedName);
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "Failed to resolve themed Nextcloud instance name for CardDAV folder.",
                    ex);
            }

            Uri uri;
            if (Uri.TryCreate(_configuration.GetNormalizedBaseUrl(), UriKind.Absolute, out uri)
                && !string.IsNullOrWhiteSpace(uri.Host))
            {
                return SanitizeFolderName(uri.Host);
            }
            return "Nextcloud";
        }

        private static string SanitizeFolderName(string value)
        {
            string result = (value ?? string.Empty).Trim();
            foreach (char c in new[] { '\\', '/', ':', '*', '?', '"', '<', '>', '|' })
            {
                result = result.Replace(c, '-');
            }
            return string.IsNullOrWhiteSpace(result) ? "Nextcloud" : result;
        }

        private XDocument SendDavRequest(string url, string method, string depth, string body)
        {
            var client = new NcHttpClient(_configuration);
            var headers = new Dictionary<string, string>();
            if (!string.IsNullOrWhiteSpace(depth))
            {
                headers["Depth"] = depth;
            }

            NcHttpResponse response = client.Send(new NcHttpRequestOptions
            {
                Method = method,
                Url = url,
                Payload = body,
                Accept = "application/xml, text/xml",
                ContentType = "application/xml; charset=utf-8",
                IncludeOcsApiHeader = false,
                ParseJson = false,
                Headers = headers
            });

            if (!response.HasHttpResponse
                || (int)response.StatusCode != 207)
            {
                throw new InvalidOperationException("CardDAV request failed for " + url + ".");
            }
            XDocument document = XDocument.Parse(response.ResponseText ?? string.Empty);
            if (document.Root == null || document.Root.Name != Dav + "multistatus"
                || document.Descendants(Dav + "error").Any()
                || document.Descendants(Dav + "status").Any(status => !IsSuccessfulStatus(status.Value)))
            {
                throw new InvalidOperationException("Incomplete or failed CardDAV response; deletion skipped.");
            }
            return document;
        }

        private static CardDavContactRecord ParseVCard(string raw)
        {
            List<string> lines = UnfoldVCard(raw);
            Dictionary<string, string> groupLabels = ReadVCardGroupLabels(lines);
            var result = new CardDavContactRecord();
            foreach (string line in lines)
            {
                int colon = line.IndexOf(':');
                if (colon <= 0)
                {
                    continue;
                }

                string left = line.Substring(0, colon);
                string upperLeft = left.ToUpperInvariant();
                string value = DecodeVCardValue(left, line.Substring(colon + 1));
                string name = NormalizeVCardPropertyName(left);

                switch (name)
                {
                    case "UID": result.Uid = value; break;
                    case "FN": result.FullName = value; break;
                    case "N":
                        string[] n = SplitVCardComponents(value);
                        result.LastName = n.Length > 0 ? n[0] : string.Empty;
                        result.FirstName = n.Length > 1 ? n[1] : string.Empty;
                        break;
                    case "ORG":
                        string[] organization = SplitVCardComponents(value);
                        result.Company = organization.Length > 0 ? organization[0] : string.Empty;
                        break;
                    case "TITLE": result.JobTitle = value; break;
                    case "EMAIL":
                        if (string.IsNullOrWhiteSpace(result.Email1)) result.Email1 = value;
                        else if (string.IsNullOrWhiteSpace(result.Email2)) result.Email2 = value;
                        break;
                    case "TEL":
                        string group = GetVCardPropertyGroup(left);
                        string groupLabel = string.Empty;
                        if (!string.IsNullOrWhiteSpace(group))
                        {
                            groupLabels.TryGetValue(group, out groupLabel);
                        }
                        AssignTelephone(result, value, upperLeft, groupLabel);
                        break;
                    case "ADR":
                        string[] address = SplitVCardComponents(value);
                        if (address.Length > 2) result.BusinessAddressStreet = address[2];
                        if (address.Length > 3) result.BusinessAddressCity = address[3];
                        if (address.Length > 4) result.BusinessAddressState = address[4];
                        if (address.Length > 5) result.BusinessAddressPostalCode = address[5];
                        if (address.Length > 6) result.BusinessAddressCountry = address[6];
                        break;
                    case "NOTE": result.Notes = value; break;
                }
            }

            if (string.IsNullOrWhiteSpace(result.FullName))
            {
                result.FullName = ((result.FirstName ?? string.Empty)
                    + " "
                    + (result.LastName ?? string.Empty)).Trim();
            }
            return result;
        }

        private static Dictionary<string, string> ReadVCardGroupLabels(IList<string> lines)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (lines == null)
            {
                return result;
            }

            foreach (string line in lines)
            {
                int colon = line.IndexOf(':');
                if (colon <= 0)
                {
                    continue;
                }

                string left = line.Substring(0, colon);
                if (!string.Equals(
                        NormalizeVCardPropertyName(left),
                        "X-ABLABEL",
                        StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string group = GetVCardPropertyGroup(left);
                if (string.IsNullOrWhiteSpace(group))
                {
                    continue;
                }

                string label = DecodeVCardValue(left, line.Substring(colon + 1));
                if (!string.IsNullOrWhiteSpace(label))
                {
                    result[group] = NormalizeAppleVCardLabel(label);
                }
            }
            return result;
        }

        private static string NormalizeVCardPropertyName(string fieldDefinition)
        {
            string token = GetVCardPropertyToken(fieldDefinition);
            int dot = token.LastIndexOf('.');
            if (dot >= 0 && dot + 1 < token.Length)
            {
                token = token.Substring(dot + 1);
            }
            return token.ToUpperInvariant();
        }

        private static string GetVCardPropertyGroup(string fieldDefinition)
        {
            string token = GetVCardPropertyToken(fieldDefinition);
            int dot = token.LastIndexOf('.');
            return dot > 0 ? token.Substring(0, dot) : string.Empty;
        }

        private static string GetVCardPropertyToken(string fieldDefinition)
        {
            if (string.IsNullOrWhiteSpace(fieldDefinition))
            {
                return string.Empty;
            }
            int semicolon = fieldDefinition.IndexOf(';');
            return (semicolon >= 0
                    ? fieldDefinition.Substring(0, semicolon)
                    : fieldDefinition)
                .Trim();
        }

        private static string NormalizeAppleVCardLabel(string value)
        {
            string label = (value ?? string.Empty).Trim();
            if (label.StartsWith("_$!<", StringComparison.Ordinal)
                && label.EndsWith(">!$_", StringComparison.Ordinal)
                && label.Length > 8)
            {
                label = label.Substring(4, label.Length - 8);
            }
            return label.Trim();
        }

        private static void AssignTelephone(
            CardDavContactRecord contact,
            string value,
            string fieldDefinition,
            string groupLabel)
        {
            if (contact == null || string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            string type = ((fieldDefinition ?? string.Empty)
                           + ";"
                           + (groupLabel ?? string.Empty))
                .ToUpperInvariant();
            string assigned;

            if (type.Contains("FAX"))
            {
                if (type.Contains("HOME"))
                {
                    if (TryAssignPhoneValue(contact.HomeFax, value, out assigned))
                    {
                        contact.HomeFax = assigned;
                    }
                }
                else if (TryAssignPhoneValue(contact.BusinessFax, value, out assigned))
                {
                    contact.BusinessFax = assigned;
                }
                return;
            }

            if (type.Contains("CELL")
                || type.Contains("MOBILE")
                || type.Contains("IPHONE"))
            {
                if (TryAssignPhoneValue(contact.MobilePhone, value, out assigned))
                {
                    contact.MobilePhone = assigned;
                }
                else if (TryAssignPhoneValue(contact.OtherPhone, value, out assigned))
                {
                    contact.OtherPhone = assigned;
                }
                return;
            }

            if (type.Contains("HOME"))
            {
                if (TryAssignPhoneValue(contact.HomePhone, value, out assigned))
                {
                    contact.HomePhone = assigned;
                }
                else if (TryAssignPhoneValue(contact.HomePhone2, value, out assigned))
                {
                    contact.HomePhone2 = assigned;
                }
                else if (TryAssignPhoneValue(contact.OtherPhone, value, out assigned))
                {
                    contact.OtherPhone = assigned;
                }
                return;
            }

            if (type.Contains("PAGER"))
            {
                if (TryAssignPhoneValue(contact.PagerPhone, value, out assigned))
                {
                    contact.PagerPhone = assigned;
                }
                return;
            }

            if (type.Contains("CAR"))
            {
                if (TryAssignPhoneValue(contact.CarPhone, value, out assigned))
                {
                    contact.CarPhone = assigned;
                }
                return;
            }

            if (type.Contains("MAIN"))
            {
                if (TryAssignPhoneValue(contact.CompanyMainPhone, value, out assigned))
                {
                    contact.CompanyMainPhone = assigned;
                }
                return;
            }

            if (TryAssignPhoneValue(contact.BusinessPhone, value, out assigned))
            {
                contact.BusinessPhone = assigned;
            }
            else if (TryAssignPhoneValue(contact.BusinessPhone2, value, out assigned))
            {
                contact.BusinessPhone2 = assigned;
            }
            else if (TryAssignPhoneValue(contact.OtherPhone, value, out assigned))
            {
                contact.OtherPhone = assigned;
            }
        }

        private static bool TryAssignPhoneValue(
            string current,
            string value,
            out string assigned)
        {
            assigned = current;
            string normalized = (value ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return false;
            }
            if (string.Equals(
                    (current ?? string.Empty).Trim(),
                    normalized,
                    StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            if (!string.IsNullOrWhiteSpace(current))
            {
                return false;
            }

            assigned = normalized;
            return true;
        }

        private static List<string> UnfoldVCard(string raw)
        {
            string normalized = (raw ?? string.Empty)
                .Replace("\r\n", "\n")
                .Replace("\r", "\n");
            string[] source = normalized.Split('\n');
            var result = new List<string>();
            foreach (string line in source)
            {
                bool foldedLine = line.StartsWith(" ", StringComparison.Ordinal)
                                  || line.StartsWith("\t", StringComparison.Ordinal);
                if (result.Count > 0
                    && result[result.Count - 1].EndsWith("=", StringComparison.Ordinal)
                    && IsQuotedPrintableProperty(result[result.Count - 1]))
                {
                    string previous = result[result.Count - 1];
                    string continuation = foldedLine && line.Length > 0
                        ? line.Substring(1)
                        : line;
                    result[result.Count - 1] =
                        previous.Substring(0, previous.Length - 1)
                        + continuation;
                }
                else if (foldedLine && result.Count > 0)
                {
                    result[result.Count - 1] += line.Substring(1);
                }
                else
                {
                    result.Add(line);
                }
            }
            return result;
        }

        private static bool IsQuotedPrintableProperty(string unfoldedLine)
        {
            if (string.IsNullOrWhiteSpace(unfoldedLine))
            {
                return false;
            }
            int colon = unfoldedLine.IndexOf(':');
            string fieldDefinition = colon >= 0
                ? unfoldedLine.Substring(0, colon)
                : unfoldedLine;
            return fieldDefinition.IndexOf(
                       "QUOTED-PRINTABLE",
                       StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string DecodeVCardValue(string fieldDefinition, string value)
        {
            string decoded = value ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(fieldDefinition)
                && fieldDefinition.IndexOf("QUOTED-PRINTABLE", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                decoded = DecodeQuotedPrintableUtf8(decoded);
            }
            return DecodeVCardText(decoded);
        }

        private static string DecodeQuotedPrintableUtf8(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            var bytes = new List<byte>();
            for (int i = 0; i < value.Length; i++)
            {
                char c = value[i];
                if (c == '=' && i + 2 < value.Length)
                {
                    int high = HexValue(value[i + 1]);
                    int low = HexValue(value[i + 2]);
                    if (high >= 0 && low >= 0)
                    {
                        bytes.Add((byte)((high << 4) | low));
                        i += 2;
                        continue;
                    }
                }

                if (c <= 0x7F)
                {
                    bytes.Add((byte)c);
                }
                else
                {
                    bytes.AddRange(Encoding.UTF8.GetBytes(new[] { c }));
                }
            }
            return Encoding.UTF8.GetString(bytes.ToArray());
        }

        private static int HexValue(char c)
        {
            if (c >= '0' && c <= '9') return c - '0';
            if (c >= 'A' && c <= 'F') return c - 'A' + 10;
            if (c >= 'a' && c <= 'f') return c - 'a' + 10;
            return -1;
        }

        private static string[] SplitVCardComponents(string value)
        {
            var parts = new List<string>();
            var current = new StringBuilder();
            bool escaped = false;
            foreach (char c in value ?? string.Empty)
            {
                if (escaped)
                {
                    current.Append('\\');
                    current.Append(c);
                    escaped = false;
                }
                else if (c == '\\')
                {
                    escaped = true;
                }
                else if (c == ';')
                {
                    parts.Add(DecodeVCardText(current.ToString()));
                    current.Clear();
                }
                else
                {
                    current.Append(c);
                }
            }
            if (escaped)
            {
                current.Append('\\');
            }
            parts.Add(DecodeVCardText(current.ToString()));
            return parts.ToArray();
        }

        private static string DecodeVCardText(string value)
        {
            return (value ?? string.Empty)
                .Replace("\\n", Environment.NewLine)
                .Replace("\\N", Environment.NewLine)
                .Replace("\\,", ",")
                .Replace("\\;", ";")
                .Replace("\\\\", "\\");
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
    }
}
