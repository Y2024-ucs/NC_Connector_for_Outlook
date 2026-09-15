// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class CardDavOutlookImportService
    {
        private const string UidPropertyName = "NC-CardDAV-UID";
        private const string HrefPropertyName = "NC-CardDAV-HREF";
        private const string ETagPropertyName = "NC-CardDAV-ETAG";
        private const string RootFolderName = "Nextcloud";
        private const string CompanyFolderName = "Firmenverzeichnis";
        private const string PersonalFolderName = "Persoenliche Kontakte";
        private const string GeneratedSystemAddressBookMarker = "z-server-generated--system";

        private readonly TalkServiceConfiguration _configuration;

        internal CardDavOutlookImportService(TalkServiceConfiguration configuration)
        {
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
        }

        internal int ImportAddressBook(
            Outlook.Application outlookApplication,
            CardDavAddressBook addressBook,
            IEnumerable<CardDavContactRecord> contacts)
        {
            if (outlookApplication == null)
            {
                throw new ArgumentNullException(nameof(outlookApplication));
            }
            if (addressBook == null)
            {
                throw new ArgumentNullException(nameof(addressBook));
            }

            Outlook.NameSpace session = null;
            Outlook.MAPIFolder defaultContacts = null;
            Outlook.MAPIFolder nextcloudFolder = null;
            Outlook.MAPIFolder instanceFolder = null;
            Outlook.MAPIFolder targetFolder = null;
            try
            {
                session = outlookApplication.Session;
                defaultContacts = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                nextcloudFolder = EnsureSubFolder(defaultContacts, RootFolderName);
                instanceFolder = EnsureSubFolder(nextcloudFolder, ResolveInstanceName());
                targetFolder = EnsureSubFolder(
                    instanceFolder,
                    IsSystemAddressBook(addressBook) ? CompanyFolderName : PersonalFolderName);

                string storeId = targetFolder.StoreID;
                Dictionary<string, string> existing = LoadExistingContactEntryIds(targetFolder);
                int imported = 0;

                foreach (CardDavContactRecord source in contacts ?? Enumerable.Empty<CardDavContactRecord>())
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
                            target = folderItems.Add(Outlook.OlItemType.olContactItem)
                                as Outlook.ContactItem;
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
                        ComInteropScope.TryRelease(target, LogCategories.Core, "Failed to release CardDAV ContactItem.");
                        ComInteropScope.TryRelease(folderItems, LogCategories.Core, "Failed to release CardDAV Items collection.");
                    }
                }

                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "CardDAV read-only sync imported/updated "
                    + imported
                    + " contacts into Outlook folder '"
                    + RootFolderName + "/" + instanceFolder.Name + "/" + targetFolder.Name + "'.");
                return imported;
            }
            finally
            {
                ComInteropScope.TryRelease(targetFolder, LogCategories.Core, "Failed to release CardDAV target folder.");
                ComInteropScope.TryRelease(instanceFolder, LogCategories.Core, "Failed to release CardDAV instance folder.");
                ComInteropScope.TryRelease(nextcloudFolder, LogCategories.Core, "Failed to release CardDAV root folder.");
                ComInteropScope.TryRelease(defaultContacts, LogCategories.Core, "Failed to release Outlook contacts folder.");
                ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session after CardDAV sync.");
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

                        string uid = ReadUserProperty(contact, UidPropertyName);
                        string href = ReadUserProperty(contact, HrefPropertyName);
                        string key = BuildContactKey(uid, href);
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
            target.MobileTelephoneNumber = source.MobilePhone ?? string.Empty;
            target.HomeTelephoneNumber = source.HomePhone ?? string.Empty;
            target.BusinessAddressStreet = source.BusinessAddressStreet ?? string.Empty;
            target.BusinessAddressCity = source.BusinessAddressCity ?? string.Empty;
            target.BusinessAddressState = source.BusinessAddressState ?? string.Empty;
            target.BusinessAddressPostalCode = source.BusinessAddressPostalCode ?? string.Empty;
            target.BusinessAddressCountry = source.BusinessAddressCountry ?? string.Empty;
            target.Body = source.Notes ?? string.Empty;

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
                return property != null ? Convert.ToString(property.Value) ?? string.Empty : string.Empty;
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
                    property = properties.Add(name, Outlook.OlUserPropertyType.olText, true, Type.Missing);
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
            return RootFolderName;
        }

        private static string SanitizeFolderName(string value)
        {
            string result = (value ?? string.Empty).Trim();
            foreach (char c in new[] { '\\', '/', ':', '*', '?', '"', '<', '>', '|' })
            {
                result = result.Replace(c, '-');
            }
            return string.IsNullOrWhiteSpace(result) ? RootFolderName : result;
        }
    }
}
