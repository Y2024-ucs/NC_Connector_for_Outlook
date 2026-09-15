// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
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
        internal string MobilePhone { get; set; }
        internal string HomePhone { get; set; }
        internal string BusinessAddressStreet { get; set; }
        internal string BusinessAddressCity { get; set; }
        internal string BusinessAddressState { get; set; }
        internal string BusinessAddressPostalCode { get; set; }
        internal string BusinessAddressCountry { get; set; }
        internal string Notes { get; set; }
    }

    internal sealed class CardDavReadOnlySync
    {
        private const string UidPropertyName = "NC-CardDAV-UID";
        private const string HrefPropertyName = "NC-CardDAV-HREF";
        private const string ETagPropertyName = "NC-CardDAV-ETAG";
        private const string RootFolderName = "Nextcloud";

        private static readonly XNamespace Dav = "DAV:";
        private static readonly XNamespace CardDav = "urn:ietf:params:xml:ns:carddav";

        private readonly TalkServiceConfiguration _configuration;

        internal CardDavReadOnlySync(TalkServiceConfiguration configuration)
        {
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
        }

        internal int Synchronize(Outlook.Application outlookApplication)
        {
            if (outlookApplication == null)
            {
                throw new ArgumentNullException(nameof(outlookApplication));
            }
            if (!_configuration.IsComplete())
            {
                throw new InvalidOperationException("Nextcloud configuration is incomplete.");
            }

            IList<CardDavAddressBook> addressBooks = new DavDiscoveryService(_configuration)
                .DiscoverAddressBooks();
            var contacts = new List<CardDavContactRecord>();
            foreach (CardDavAddressBook addressBook in addressBooks)
            {
                contacts.AddRange(DownloadContacts(addressBook));
            }

            return ImportIntoOutlook(outlookApplication, contacts);
        }

        internal IList<CardDavContactRecord> DownloadContacts(CardDavAddressBook addressBook)
        {
            if (addressBook == null || string.IsNullOrWhiteSpace(addressBook.Href))
            {
                throw new ArgumentException("A CardDAV address book is required.", nameof(addressBook));
            }

            string body = "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                        + "<card:addressbook-query xmlns:d=\"DAV:\" xmlns:card=\"urn:ietf:params:xml:ns:carddav\">"
                        + "<d:prop><d:getetag/><card:address-data content-type=\"text/vcard\"/></d:prop>"
                        + "<card:filter><card:prop-filter name=\"FN\"/></card:filter>"
                        + "</card:addressbook-query>";

            XDocument document = SendDavRequest(addressBook.Href, "REPORT", "1", body);
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

                string vcard = (string)prop.Element(CardDav + "address-data");
                if (string.IsNullOrWhiteSpace(vcard))
                {
                    continue;
                }

                CardDavContactRecord contact = ParseVCard(vcard);
                contact.Href = ResolveUri(
                    addressBook.Href,
                    (string)responseElement.Element(Dav + "href"));
                contact.ETag = ((string)prop.Element(Dav + "getetag") ?? string.Empty).Trim();
                contacts.Add(contact);
            }

            return contacts;
        }

        internal int ImportIntoOutlook(
            Outlook.Application outlookApplication,
            IEnumerable<CardDavContactRecord> contacts)
        {
            if (outlookApplication == null)
            {
                throw new ArgumentNullException(nameof(outlookApplication));
            }

            Outlook.NameSpace session = null;
            Outlook.MAPIFolder defaultContacts = null;
            Outlook.MAPIFolder nextcloudFolder = null;
            Outlook.MAPIFolder instanceFolder = null;
            try
            {
                session = outlookApplication.Session;
                defaultContacts = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                nextcloudFolder = EnsureSubFolder(defaultContacts, RootFolderName);
                instanceFolder = EnsureSubFolder(nextcloudFolder, ResolveInstanceName());

                string storeId = instanceFolder.StoreID;
                Dictionary<string, string> existing = LoadExistingContactEntryIds(instanceFolder);
                int imported = 0;

                foreach (CardDavContactRecord source in contacts ?? Enumerable.Empty<CardDavContactRecord>())
                {
                    if (source == null)
                    {
                        continue;
                    }

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
                            folderItems = instanceFolder.Items;
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
                        ComInteropScope.TryRelease(
                            target,
                            LogCategories.Core,
                            "Failed to release CardDAV ContactItem.");
                        ComInteropScope.TryRelease(
                            folderItems,
                            LogCategories.Core,
                            "Failed to release CardDAV Items collection.");
                    }
                }

                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "CardDAV read-only sync imported/updated " + imported + " contacts into Outlook folder '"
                    + RootFolderName + "/" + instanceFolder.Name + "'.");
                return imported;
            }
            finally
            {
                ComInteropScope.TryRelease(
                    instanceFolder,
                    LogCategories.Core,
                    "Failed to release CardDAV instance folder.");
                ComInteropScope.TryRelease(
                    nextcloudFolder,
                    LogCategories.Core,
                    "Failed to release CardDAV root folder.");
                ComInteropScope.TryRelease(
                    defaultContacts,
                    LogCategories.Core,
                    "Failed to release Outlook contacts folder.");
                ComInteropScope.TryRelease(
                    session,
                    LogCategories.Core,
                    "Failed to release Outlook session after CardDAV sync.");
            }
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
                            ComInteropScope.TryRelease(
                                contact,
                                LogCategories.Core,
                                "Failed to release existing CardDAV ContactItem.");
                        }
                        else
                        {
                            ComInteropScope.TryRelease(
                                raw,
                                LogCategories.Core,
                                "Failed to release non-contact CardDAV folder item.");
                        }
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(
                    items,
                    LogCategories.Core,
                    "Failed to release CardDAV contact items collection.");
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

        private static string BuildContactKey(CardDavContactRecord contact)
        {
            if (contact == null)
            {
                return string.Empty;
            }
            return BuildContactKey(contact.Uid, contact.Href);
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
                ComInteropScope.TryRelease(
                    property,
                    LogCategories.Core,
                    "Failed to release CardDAV contact user property.");
                ComInteropScope.TryRelease(
                    properties,
                    LogCategories.Core,
                    "Failed to release CardDAV contact user properties.");
            }
        }

        private static void WriteUserProperty(
            Outlook.ContactItem item,
            string name,
            string value)
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
                ComInteropScope.TryRelease(
                    property,
                    LogCategories.Core,
                    "Failed to release CardDAV contact user property.");
                ComInteropScope.TryRelease(
                    properties,
                    LogCategories.Core,
                    "Failed to release CardDAV contact user properties.");
            }
        }

        private static Outlook.MAPIFolder EnsureSubFolder(
            Outlook.MAPIFolder parent,
            string name)
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
                            && string.Equals(
                                folder.Name,
                                name,
                                StringComparison.OrdinalIgnoreCase))
                        {
                            Outlook.MAPIFolder match = folder;
                            folder = null;
                            return match;
                        }
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(
                            folder,
                            LogCategories.Core,
                            "Failed to release Outlook contact subfolder.");
                    }
                }

                return folders.Add(
                    name,
                    Outlook.OlDefaultFolders.olFolderContacts);
            }
            finally
            {
                ComInteropScope.TryRelease(
                    folders,
                    LogCategories.Core,
                    "Failed to release Outlook contact folders collection.");
            }
        }

        private string ResolveInstanceName()
        {
            try
            {
                NextcloudCapabilitiesSnapshot snapshot =
                    new NextcloudCapabilitiesService(_configuration)
                        .GetSnapshot(false, false);
                IDictionary<string, object> theming = NcJson.GetDictionary(
                    snapshot != null ? snapshot.Capabilities : null,
                    "theming");
                string themedName = NcJson.GetTrimmedString(theming, "name");
                if (!string.IsNullOrWhiteSpace(themedName)
                    && !string.Equals(
                        themedName,
                        "Nextcloud",
                        StringComparison.OrdinalIgnoreCase))
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
            if (Uri.TryCreate(
                    _configuration.GetNormalizedBaseUrl(),
                    UriKind.Absolute,
                    out uri)
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

        private XDocument SendDavRequest(
            string url,
            string method,
            string depth,
            string body)
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
                || (int)response.StatusCode < 200
                || (int)response.StatusCode >= 300)
            {
                throw new InvalidOperationException(
                    "CardDAV request failed for " + url + ".");
            }
            return XDocument.Parse(response.ResponseText ?? string.Empty);
        }

        private static CardDavContactRecord ParseVCard(string raw)
        {
            List<string> lines = UnfoldVCard(raw);
            var result = new CardDavContactRecord();
            foreach (string line in lines)
            {
                int colon = line.IndexOf(':');
                if (colon <= 0)
                {
                    continue;
                }

                string left = line.Substring(0, colon);
                string value = DecodeVCardText(line.Substring(colon + 1));
                string name = left.Split(';')[0].ToUpperInvariant();
                string upperLeft = left.ToUpperInvariant();

                switch (name)
                {
                    case "UID":
                        result.Uid = value;
                        break;
                    case "FN":
                        result.FullName = value;
                        break;
                    case "N":
                        string[] n = SplitVCardComponents(value);
                        result.LastName = n.Length > 0 ? n[0] : string.Empty;
                        result.FirstName = n.Length > 1 ? n[1] : string.Empty;
                        break;
                    case "ORG":
                        string[] organization = SplitVCardComponents(value);
                        result.Company = organization.Length > 0
                            ? organization[0]
                            : string.Empty;
                        break;
                    case "TITLE":
                        result.JobTitle = value;
                        break;
                    case "EMAIL":
                        if (string.IsNullOrWhiteSpace(result.Email1))
                        {
                            result.Email1 = value;
                        }
                        else if (string.IsNullOrWhiteSpace(result.Email2))
                        {
                            result.Email2 = value;
                        }
                        break;
                    case "TEL":
                        if (upperLeft.Contains("CELL") || upperLeft.Contains("MOBILE"))
                        {
                            result.MobilePhone = value;
                        }
                        else if (upperLeft.Contains("HOME"))
                        {
                            result.HomePhone = value;
                        }
                        else if (string.IsNullOrWhiteSpace(result.BusinessPhone))
                        {
                            result.BusinessPhone = value;
                        }
                        break;
                    case "ADR":
                        string[] address = SplitVCardComponents(value);
                        if (address.Length > 2) result.BusinessAddressStreet = address[2];
                        if (address.Length > 3) result.BusinessAddressCity = address[3];
                        if (address.Length > 4) result.BusinessAddressState = address[4];
                        if (address.Length > 5) result.BusinessAddressPostalCode = address[5];
                        if (address.Length > 6) result.BusinessAddressCountry = address[6];
                        break;
                    case "NOTE":
                        result.Notes = value;
                        break;
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

        private static List<string> UnfoldVCard(string raw)
        {
            string normalized = (raw ?? string.Empty)
                .Replace("\r\n", "\n")
                .Replace("\r", "\n");
            string[] source = normalized.Split('\n');
            var result = new List<string>();
            foreach (string line in source)
            {
                if ((line.StartsWith(" ") || line.StartsWith("\t"))
                    && result.Count > 0)
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

        private static string[] SplitVCardComponents(string value)
        {
            var parts = new List<string>();
            var current = new System.Text.StringBuilder();
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
