// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Reflection;
using System.Threading;
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

    internal static class CardDavBackgroundSyncManager
    {
        private static readonly object SyncRoot = new object();
        private static readonly TimeSpan Interval = TimeSpan.FromMinutes(15);
        private static Timer _timer;
        private static TalkServiceConfiguration _configuration;
        private static Outlook.Application _outlookApplication;
        private static OutlookUiSynchronizationContext _uiContext;
        private static int _running;

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
            if (Interlocked.CompareExchange(ref _running, 1, 0) != 0)
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
                            DiagnosticsLogger.Log(LogCategories.Core, "CardDAV background sync completed (contacts=" + count + ").");
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
