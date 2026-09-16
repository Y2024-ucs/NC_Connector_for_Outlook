// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Xml.Linq;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Services
{
    internal static class CardDavDefaultContactsImporter
    {
        private const string UidPropertyName = "NC-CardDAV-UID";
        private const string HrefPropertyName = "NC-CardDAV-HREF";
        private const string ETagPropertyName = "NC-CardDAV-ETAG";

        internal static int Import(
            Outlook.Application outlookApplication,
            IEnumerable<CardDavContactRecord> contacts)
        {
            if (outlookApplication == null)
            {
                throw new ArgumentNullException("outlookApplication");
            }

            Outlook.NameSpace session = null;
            Outlook.MAPIFolder defaultContacts = null;
            try
            {
                session = outlookApplication.Session;
                defaultContacts = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                return ImportIntoFolder(session, defaultContacts, contacts);
            }
            finally
            {
                ComInteropScope.TryRelease(defaultContacts, LogCategories.Core, "Failed to release default Outlook contacts folder.");
                ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session after default CardDAV import.");
            }
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
                if (source == null)
                {
                    continue;
                }

                string key = BuildContactKey(source.Uid, source.Href);
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
                        target = folderItems.Add(Outlook.OlItemType.olContactItem) as Outlook.ContactItem;
                    }
                    if (target == null)
                    {
                        continue;
                    }

                    ApplyContact(target, source);
                    target.Save();
                    CardDavReadOnlySync.RememberKnownEtag(source.Href, source.ETag);
                    imported++;
                }
                finally
                {
                    ComInteropScope.TryRelease(target, LogCategories.Core, "Failed to release default CardDAV ContactItem.");
                    ComInteropScope.TryRelease(folderItems, LogCategories.Core, "Failed to release default CardDAV Items collection.");
                }
            }

            DiagnosticsLogger.Log(
                LogCategories.Core,
                "CardDAV read-only sync imported/updated " + imported + " contacts in the default Outlook contacts folder.");
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
                            ComInteropScope.TryRelease(contact, LogCategories.Core, "Failed to release existing default CardDAV ContactItem.");
                        }
                        else
                        {
                            ComInteropScope.TryRelease(raw, LogCategories.Core, "Failed to release non-contact item in default contacts folder.");
                        }
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(items, LogCategories.Core, "Failed to release default contacts Items collection.");
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
            target.MobileTelephoneNumber = source.MobilePhone ?? string.Empty;
            target.HomeTelephoneNumber = source.HomePhone ?? string.Empty;
            target.BusinessAddressStreet = source.BusinessAddressStreet ?? string.Empty;
            target.BusinessAddressCity = source.BusinessAddressCity ?? string.Empty;
            target.BusinessAddressState = source.BusinessAddressState ?? string.Empty;
            target.BusinessAddressPostalCode = source.BusinessAddressPostalCode ?? string.Empty;
            target.BusinessAddressCountry = source.BusinessAddressCountry ?? string.Empty;
            target.Body = source.Notes ?? string.Empty;
            CardDavVCardSupport.ApplyPhoto(target, source.Href);

            WriteUserProperty(target, UidPropertyName, source.Uid);
            WriteUserProperty(target, HrefPropertyName, source.Href);
            WriteUserProperty(target, ETagPropertyName, source.ETag);
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
                ComInteropScope.TryRelease(property, LogCategories.Core, "Failed to release default CardDAV contact user property.");
                ComInteropScope.TryRelease(properties, LogCategories.Core, "Failed to release default CardDAV contact user properties.");
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
                ComInteropScope.TryRelease(property, LogCategories.Core, "Failed to release default CardDAV contact user property.");
                ComInteropScope.TryRelease(properties, LogCategories.Core, "Failed to release default CardDAV contact user properties.");
            }
        }
    }

    internal static class CardDavContactGroupSync
    {
        private const string HrefPropertyName = "NC-CardDAV-HREF";
        private const string ManagedGroupPropertyName = "NC-CardDAV-GROUP";
        private const string ManagedGroupNamePropertyName = "NC-CardDAV-GROUP-NAME";
        private const string GeneratedSystemAddressBookMarker = "z-server-generated--system";
        private static readonly XNamespace Dav = "DAV:";
        private static readonly XNamespace CardDav = "urn:ietf:params:xml:ns:carddav";

        private sealed class GroupMember
        {
            internal string DisplayName { get; set; }
            internal string EmailAddress { get; set; }
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

            Outlook.NameSpace session = null;
            Outlook.MAPIFolder defaultContacts = null;
            Outlook.Folders folders = null;
            int groupCount = 0;
            try
            {
                session = outlookApplication.Session;
                defaultContacts = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                groupCount += ReconcileFolder(session, defaultContacts, categoriesByHref);

                folders = defaultContacts.Folders;
                int count = folders != null ? folders.Count : 0;
                for (int index = 1; index <= count; index++)
                {
                    Outlook.MAPIFolder folder = null;
                    try
                    {
                        folder = folders[index];
                        if (folder != null)
                        {
                            groupCount += ReconcileFolder(session, folder, categoriesByHref);
                        }
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(folder, LogCategories.Core, "Failed to release CardDAV group contact subfolder.");
                    }
                }

                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "CardDAV contact group reconciliation completed (groups=" + groupCount + ").");
                return groupCount;
            }
            finally
            {
                ComInteropScope.TryRelease(folders, LogCategories.Core, "Failed to release CardDAV group contact subfolders.");
                ComInteropScope.TryRelease(defaultContacts, LogCategories.Core, "Failed to release default contacts folder after CardDAV group sync.");
                ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session after CardDAV group sync.");
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
                    XElement prop = null;
                    foreach (XElement propstat in responseElement.Elements(Dav + "propstat"))
                    {
                        XElement status = propstat.Element(Dav + "status");
                        if (status != null && IsSuccessfulStatus(status.Value))
                        {
                            prop = propstat.Element(Dav + "prop");
                            if (prop != null)
                            {
                                break;
                            }
                        }
                    }
                    if (prop == null)
                    {
                        continue;
                    }

                    string requestHref = ((string)responseElement.Element(Dav + "href") ?? string.Empty).Trim();
                    string href = ResolveUri(addressBook.Href, requestHref);
                    string vcard = (string)prop.Element(CardDav + "address-data");
                    if (string.IsNullOrWhiteSpace(href) || string.IsNullOrWhiteSpace(vcard))
                    {
                        continue;
                    }
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
                throw new InvalidOperationException("CardDAV address book is missing for group sync.");
            }

            string body = "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                + "<card:addressbook-query xmlns:d=\"DAV:\" xmlns:card=\"urn:ietf:params:xml:ns:carddav\">"
                + "<d:prop><card:address-data content-type=\"text/vcard\">"
                + "<card:prop name=\"UID\"/><card:prop name=\"FN\"/><card:prop name=\"EMAIL\"/><card:prop name=\"CATEGORIES\"/>"
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
                throw new InvalidOperationException("CardDAV group query failed for " + addressBook.Href + ".");
            }

            XDocument document = XDocument.Parse(response.ResponseText ?? string.Empty);
            if (document.Root == null || document.Root.Name != Dav + "multistatus")
            {
                throw new InvalidOperationException("Invalid CardDAV group response for " + addressBook.Href + ".");
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
                if (colon <= 0)
                {
                    continue;
                }
                string field = line.Substring(0, colon);
                string name = field.Split(';')[0];
                if (!string.Equals(name, "CATEGORIES", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                foreach (string category in SplitCategoryValues(line.Substring(colon + 1)))
                {
                    string value = (category ?? string.Empty).Trim();
                    if (!string.IsNullOrWhiteSpace(value) && !result.Contains(value))
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
                    if (c == 'n' || c == 'N')
                    {
                        current.Append(Environment.NewLine);
                    }
                    else
                    {
                        current.Append(c);
                    }
                    escaped = false;
                }
                else if (c == '\\')
                {
                    escaped = true;
                }
                else if (c == ',')
                {
                    result.Add(current.ToString());
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
            result.Add(current.ToString());
            return result;
        }

        private static int ReconcileFolder(
            Outlook.NameSpace session,
            Outlook.MAPIFolder folder,
            IDictionary<string, List<string>> categoriesByHref)
        {
            Dictionary<string, List<GroupMember>> groups = LoadManagedContactGroups(folder, categoriesByHref);
            DeleteManagedGroups(folder);

            int created = 0;
            foreach (KeyValuePair<string, List<GroupMember>> pair in groups)
            {
                Outlook.Items items = null;
                Outlook.DistListItem list = null;
                try
                {
                    items = folder.Items;
                    list = items.Add(Outlook.OlItemType.olDistributionListItem) as Outlook.DistListItem;
                    if (list == null)
                    {
                        continue;
                    }

                    list.DLName = pair.Key + " (Verteiler)";
                    WriteGroupProperty(list, ManagedGroupPropertyName, "1");
                    WriteGroupProperty(list, ManagedGroupNamePropertyName, pair.Key);

                    int addedMembers = 0;
                    foreach (GroupMember member in pair.Value)
                    {
                        Outlook.Recipient recipient = null;
                        try
                        {
                            string address = !string.IsNullOrWhiteSpace(member.EmailAddress)
                                ? member.EmailAddress
                                : member.DisplayName;
                            if (string.IsNullOrWhiteSpace(address))
                            {
                                continue;
                            }

                            recipient = session.CreateRecipient(address);
                            if (recipient == null || !recipient.Resolve())
                            {
                                DiagnosticsLogger.Log(
                                    LogCategories.Core,
                                    "CardDAV group member unresolved (group=" + pair.Key
                                    + ", member=" + (member.DisplayName ?? string.Empty) + ").");
                                continue;
                            }

                            list.AddMember(recipient);
                            addedMembers++;
                        }
                        finally
                        {
                            ComInteropScope.TryRelease(recipient, LogCategories.Core, "Failed to release CardDAV group recipient.");
                        }
                    }

                    if (addedMembers > 0)
                    {
                        list.Save();
                        created++;
                    }
                    else
                    {
                        list.Delete();
                    }
                }
                finally
                {
                    ComInteropScope.TryRelease(list, LogCategories.Core, "Failed to release CardDAV distribution list.");
                    ComInteropScope.TryRelease(items, LogCategories.Core, "Failed to release CardDAV distribution list items collection.");
                }
            }
            return created;
        }

        private static Dictionary<string, List<GroupMember>> LoadManagedContactGroups(
            Outlook.MAPIFolder folder,
            IDictionary<string, List<string>> categoriesByHref)
        {
            var groups = new Dictionary<string, List<GroupMember>>(StringComparer.OrdinalIgnoreCase);
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

                        string href = ReadContactProperty(contact, HrefPropertyName);
                        List<string> categories;
                        if (string.IsNullOrWhiteSpace(href)
                            || !categoriesByHref.TryGetValue(href, out categories)
                            || categories == null)
                        {
                            continue;
                        }

                        foreach (string rawCategory in categories)
                        {
                            string category = (rawCategory ?? string.Empty).Trim();
                            if (string.IsNullOrWhiteSpace(category))
                            {
                                continue;
                            }

                            List<GroupMember> members;
                            if (!groups.TryGetValue(category, out members))
                            {
                                members = new List<GroupMember>();
                                groups[category] = members;
                            }
                            members.Add(new GroupMember
                            {
                                DisplayName = contact.FullName ?? string.Empty,
                                EmailAddress = contact.Email1Address ?? string.Empty
                            });
                        }
                    }
                    finally
                    {
                        if (contact != null)
                        {
                            ComInteropScope.TryRelease(contact, LogCategories.Core, "Failed to release CardDAV group source contact.");
                        }
                        else
                        {
                            ComInteropScope.TryRelease(raw, LogCategories.Core, "Failed to release non-contact during CardDAV group scan.");
                        }
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(items, LogCategories.Core, "Failed to release contacts collection during CardDAV group scan.");
            }
            return groups;
        }

        private static void DeleteManagedGroups(Outlook.MAPIFolder folder)
        {
            Outlook.Items items = null;
            try
            {
                items = folder.Items;
                for (int index = items.Count; index >= 1; index--)
                {
                    object raw = null;
                    Outlook.DistListItem list = null;
                    try
                    {
                        raw = items[index];
                        list = raw as Outlook.DistListItem;
                        if (list != null
                            && string.Equals(ReadGroupProperty(list, ManagedGroupPropertyName), "1", StringComparison.Ordinal))
                        {
                            list.Delete();
                        }
                    }
                    finally
                    {
                        if (list != null)
                        {
                            ComInteropScope.TryRelease(list, LogCategories.Core, "Failed to release managed CardDAV distribution list during deletion.");
                        }
                        else
                        {
                            ComInteropScope.TryRelease(raw, LogCategories.Core, "Failed to release non-distribution-list during CardDAV group cleanup.");
                        }
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(items, LogCategories.Core, "Failed to release CardDAV group cleanup items collection.");
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
                ComInteropScope.TryRelease(property, LogCategories.Core, "Failed to release CardDAV contact group property.");
                ComInteropScope.TryRelease(properties, LogCategories.Core, "Failed to release CardDAV contact group properties.");
            }
        }

        private static string ReadGroupProperty(Outlook.DistListItem item, string name)
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
                ComInteropScope.TryRelease(property, LogCategories.Core, "Failed to release CardDAV group marker property.");
                ComInteropScope.TryRelease(properties, LogCategories.Core, "Failed to release CardDAV group marker properties.");
            }
        }

        private static void WriteGroupProperty(Outlook.DistListItem item, string name, string value)
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
                ComInteropScope.TryRelease(property, LogCategories.Core, "Failed to release CardDAV group marker property after write.");
                ComInteropScope.TryRelease(properties, LogCategories.Core, "Failed to release CardDAV group marker properties after write.");
            }
        }

        private static bool IsSystemAddressBook(CardDavAddressBook addressBook)
        {
            return addressBook != null
                && !string.IsNullOrWhiteSpace(addressBook.Href)
                && addressBook.Href.IndexOf(
                    GeneratedSystemAddressBookMarker,
                    StringComparison.OrdinalIgnoreCase) >= 0;
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

    internal static class CardDavBackgroundSyncManager
    {
        private static readonly object SyncRoot = new object();
        private static readonly TimeSpan Interval = TimeSpan.FromMinutes(15);
        private static Timer _timer;
        private static TalkServiceConfiguration _configuration;
        private static Outlook.Application _outlookApplication;
        private static OutlookUiSynchronizationContext _uiContext;
        private static int _running;

        internal static bool TryBeginSync()
        {
            return Interlocked.CompareExchange(ref _running, 1, 0) == 0;
        }

        internal static void EndSync()
        {
            Interlocked.Exchange(ref _running, 0);
        }

        internal static void EnsureStarted(
            TalkServiceConfiguration configuration,
            Outlook.Application outlookApplication)
        {
            if (configuration == null || outlookApplication == null)
            {
                return;
            }

            lock (SyncRoot)
            {
                _configuration = configuration;
                _outlookApplication = outlookApplication;

                if (_uiContext == null)
                {
                    _uiContext = new OutlookUiSynchronizationContext();
                }

                if (_timer != null)
                {
                    return;
                }

                _timer = new Timer(OnTimer, null, Interval, Interval);
                DiagnosticsLogger.Log(LogCategories.Core, "CardDAV background sync enabled (intervalMinutes=15).");
            }
        }

        private static void OnTimer(object state)
        {
            if (!TryBeginSync())
            {
                return;
            }

            try
            {
                CardDavSyncPreferences preferences = CardDavSyncPreferences.Load();
                if (!preferences.Enabled
                    || (!preferences.SyncCompanyDirectory && !preferences.SyncPersonalContacts))
                {
                    Interlocked.Exchange(ref _running, 0);
                    return;
                }

                TalkServiceConfiguration configuration;
                Outlook.Application outlookApplication;
                OutlookUiSynchronizationContext uiContext;
                lock (SyncRoot)
                {
                    configuration = _configuration;
                    outlookApplication = _outlookApplication;
                    uiContext = _uiContext;
                }

                if (configuration == null
                    || !configuration.IsComplete()
                    || outlookApplication == null
                    || uiContext == null)
                {
                    Interlocked.Exchange(ref _running, 0);
                    return;
                }

                var sync = new CardDavReadOnlySync(configuration);
                IList<CardDavAddressBook> addressBooks = new DavDiscoveryService(configuration).DiscoverAddressBooks();
                var contacts = new List<CardDavContactRecord>();

                foreach (CardDavAddressBook addressBook in addressBooks)
                {
                    bool isSystem = IsSystemAddressBook(addressBook);
                    if ((isSystem && !preferences.SyncCompanyDirectory)
                        || (!isSystem && !preferences.SyncPersonalContacts))
                    {
                        continue;
                    }
                    contacts.AddRange(sync.DownloadContacts(addressBook));
                }

                uiContext.Post(
                    delegate
                    {
                        try
                        {
                            int count = sync.ImportIntoOutlook(outlookApplication, contacts);
                            int groups = CardDavContactGroupSync.Reconcile(
                                configuration,
                                outlookApplication,
                                addressBooks,
                                preferences);
                            DiagnosticsLogger.Log(LogCategories.Core, "CardDAV background sync completed (contacts=" + count
                                + ", deleted=" + sync.DeletedCount + ", groups=" + groups + ").");
                        }
                        catch (Exception ex)
                        {
                            DiagnosticsLogger.LogException(LogCategories.Core, "CardDAV background Outlook import failed.", ex);
                        }
                        finally
                        {
                            Interlocked.Exchange(ref _running, 0);
                        }
                    },
                    null);
            }
            catch (Exception ex)
            {
                Interlocked.Exchange(ref _running, 0);
                DiagnosticsLogger.LogException(LogCategories.Core, "CardDAV background sync failed.", ex);
            }
        }

        private static bool IsSystemAddressBook(CardDavAddressBook addressBook)
        {
            return addressBook != null
                && !string.IsNullOrWhiteSpace(addressBook.Href)
                && addressBook.Href.IndexOf(
                    "z-server-generated--system",
                    StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
