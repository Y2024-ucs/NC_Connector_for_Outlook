// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Interop.Outlook;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

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
        internal string Email1 { get; set; }
        internal string Email2 { get; set; }
        internal string BusinessPhone { get; set; }
        internal string MobilePhone { get; set; }
        internal string HomePhone { get; set; }
        internal string BusinessAddressStreet { get; set; }
        internal string BusinessAddressCity { get; set; }
        internal string BusinessAddressPostalCode { get; set; }
        internal string BusinessAddressCountry { get; set; }
        internal string Notes { get; set; }
    }

    internal sealed class CardDavReadOnlySync
    {
        private const string UidPropertyName = "NC-CardDAV-UID";
        private const string HrefPropertyName = "NC-CardDAV-HREF";
        private const string ETagPropertyName = "NC-CardDAV-ETAG";
        private static readonly XNamespace Dav = "DAV:";
        private static readonly XNamespace CardDav = "urn:ietf:params:xml:ns:carddav";

        private readonly TalkServiceConfiguration _configuration;

        internal CardDavReadOnlySync(TalkServiceConfiguration configuration)
        {
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
        }

        internal IList<CardDavContactRecord> DownloadContacts(CardDavAddressBook addressBook)
        {
            if (addressBook == null || string.IsNullOrWhiteSpace(addressBook.Href))
            {
                throw new ArgumentException("A CardDAV address book is required.", nameof(addressBook));
            }

            string body = "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                        + "<card:addressbook-query xmlns:d=\"DAV:\" xmlns:card=\"urn:ietf:params:xml:ns:carddav\">"
                        + "<d:prop><d:getetag/><card:address-data/></d:prop>"
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
                contact.Href = ResolveUri(addressBook.Href, (string)responseElement.Element(Dav + "href"));
                contact.ETag = ((string)prop.Element(Dav + "getetag") ?? string.Empty).Trim();
                contacts.Add(contact);
            }

            return contacts;
        }

        internal int ImportIntoOutlook(Application outlookApplication, IEnumerable<CardDavContactRecord> contacts)
        {
            if (outlookApplication == null)
            {
                throw new ArgumentNullException(nameof(outlookApplication));
            }

            MAPIFolder defaultContacts = null;
            MAPIFolder nextcloudFolder = null;
            MAPIFolder instanceFolder = null;
            NameSpace session = null;
            try
            {
                session = outlookApplication.Session;
                defaultContacts = session.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
                nextcloudFolder = EnsureSubFolder(defaultContacts, "Nextcloud");
                instanceFolder = EnsureSubFolder(nextcloudFolder, ResolveInstanceName());

                Dictionary<string, ContactItem> existing = LoadExistingContacts(instanceFolder);
                int imported = 0;
                foreach (CardDavContactRecord source in contacts ?? Enumerable.Empty<CardDavContactRecord>())
                {
                    string key = BuildContactKey(source);
                    ContactItem target = null;
                    try
                    {
                        if (!existing.TryGetValue(key, out target) || target == null)
                        {
                            target = (ContactItem)instanceFolder.Items.Add(OlItemType.olContactItem);
                        }

                        ApplyContact(target, source);
                        target.Save();
                        imported++;
                    }
                    finally
                    {
                        if (target != null)
                        {
                            ComInteropScope.TryRelease(target, LogCategories.Core, "Failed to release CardDAV ContactItem.");
                        }
                    }
                }

                return imported;
            }
            finally
            {
                ReleaseContactDictionary(null);
                ComInteropScope.TryRelease(instanceFolder, LogCategories.Core, "Failed to release CardDAV instance folder.");
                ComInteropScope.TryRelease(nextcloudFolder, LogCategories.Core, "Failed to release CardDAV root folder.");
                ComInteropScope.TryRelease(defaultContacts, LogCategories.Core, "Failed to release Outlook contacts folder.");
                ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session after CardDAV sync.");
            }
        }

        private Dictionary<string, ContactItem> LoadExistingContacts(MAPIFolder folder)
        {
            var result = new Dictionary<string, ContactItem>(StringComparer.OrdinalIgnoreCase);
            Items items = null;
            try
            {
                items = folder.Items;
                foreach (object raw in items)
                {
                    ContactItem contact = raw as ContactItem;
                    if (contact == null)
                    {
                        ComInteropScope.TryRelease(raw, LogCategories.Core, "Failed to release non-contact CardDAV folder item.");
                        continue;
                    }

                    string uid = ReadUserProperty(contact, UidPropertyName);
                    string href = ReadUserProperty(contact, HrefPropertyName);
                    string key = !string.IsNullOrWhiteSpace(uid) ? "uid:" + uid : "href:" + href;
                    if (!string.IsNullOrWhiteSpace(uid) || !string.IsNullOrWhiteSpace(href))
                    {
                        if (!result.ContainsKey(key))
                        {
                            result[key] = contact;
                            continue;
                        }
                    }

                    ComInteropScope.TryRelease(contact, LogCategories.Core, "Failed to release unmatched CardDAV contact.");
                }
            }
            finally
            {
                ComInteropScope.TryRelease(items, LogCategories.Core, "Failed to release CardDAV contact items collection.");
            }

            return result;
        }

        private static void ReleaseContactDictionary(Dictionary<string, ContactItem> contacts)
        {
            if (contacts == null)
            {
                return;
            }

            foreach (ContactItem item in contacts.Values.Distinct())
            {
                ComInteropScope.TryRelease(item, LogCategories.Core, "Failed to release cached CardDAV ContactItem.");
            }
            contacts.Clear();
        }

        private static void ApplyContact(ContactItem target, CardDavContactRecord source)
        {
            target.FullName = source.FullName ?? string.Empty;
            target.FirstName = source.FirstName ?? string.Empty;
            target.LastName = source.LastName ?? string.Empty;
            target.CompanyName = source.Company ?? string.Empty;
            target.Email1Address = source.Email1 ?? string.Empty;
            target.Email2Address = source.Email2 ?? string.Empty;
            target.BusinessTelephoneNumber = source.BusinessPhone ?? string.Empty;
            target.MobileTelephoneNumber = source.MobilePhone ?? string.Empty;
            target.HomeTelephoneNumber = source.HomePhone ?? string.Empty;
            target.BusinessAddressStreet = source.BusinessAddressStreet ?? string.Empty;
            target.BusinessAddressCity = source.BusinessAddressCity ?? string.Empty;
            target.BusinessAddressPostalCode = source.BusinessAddressPostalCode ?? string.Empty;
            target.BusinessAddressCountry = source.BusinessAddressCountry ?? string.Empty;
            target.Body = source.Notes ?? string.Empty;

            WriteUserProperty(target, UidPropertyName, source.Uid);
            WriteUserProperty(target, HrefPropertyName, source.Href);
            WriteUserProperty(target, ETagPropertyName, source.ETag);
        }

        private static string BuildContactKey(CardDavContactRecord contact)
        {
            return !string.IsNullOrWhiteSpace(contact.Uid)
                ? "uid:" + contact.Uid.Trim()
                : "href:" + (contact.Href ?? string.Empty).Trim();
        }

        private static string ReadUserProperty(ContactItem item, string name)
        {
            UserProperties properties = null;
            UserProperty property = null;
            try
            {
                properties = item.UserProperties;
                property = properties.Find(name, true);
                return property != null ? Convert.ToString(property.Value) ?? string.Empty : string.Empty;
            }
            finally
            {
                ComInteropScope.TryRelease(property, LogCategories.Core, "Failed to release CardDAV contact user property.");
                ComInteropScope.TryRelease(properties, LogCategories.Core, "Failed to release CardDAV contact user properties.");
            }
        }

        private static void WriteUserProperty(ContactItem item, string name, string value)
        {
            UserProperties properties = null;
            UserProperty property = null;
            try
            {
                properties = item.UserProperties;
                property = properties.Find(name, true) ?? properties.Add(name, OlUserPropertyType.olText, true, OlFormatText.olFormatTextText);
                property.Value = value ?? string.Empty;
            }
            finally
            {
                ComInteropScope.TryRelease(property, LogCategories.Core, "Failed to release CardDAV contact user property.");
                ComInteropScope.TryRelease(properties, LogCategories.Core, "Failed to release CardDAV contact user properties.");
            }
        }

        private static MAPIFolder EnsureSubFolder(MAPIFolder parent, string name)
        {
            Folders folders = null;
            try
            {
                folders = parent.Folders;
                foreach (MAPIFolder folder in folders)
                {
                    if (string.Equals(folder.Name, name, StringComparison.OrdinalIgnoreCase))
                    {
                        return folder;
                    }
                    ComInteropScope.TryRelease(folder, LogCategories.Core, "Failed to release Outlook contact subfolder.");
                }

                return folders.Add(name, OlDefaultFolders.olFolderContacts);
            }
            finally
            {
                ComInteropScope.TryRelease(folders, LogCategories.Core, "Failed to release Outlook contact folders collection.");
            }
        }

        private string ResolveInstanceName()
        {
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

            if (!response.HasHttpResponse || ((int)response.StatusCode < 200 || (int)response.StatusCode >= 300))
            {
                throw new InvalidOperationException("CardDAV request failed for " + url + ".");
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
                    case "UID": result.Uid = value; break;
                    case "FN": result.FullName = value; break;
                    case "N":
                        string[] n = value.Split(';');
                        result.LastName = n.Length > 0 ? n[0] : string.Empty;
                        result.FirstName = n.Length > 1 ? n[1] : string.Empty;
                        break;
                    case "ORG": result.Company = value.Split(';')[0]; break;
                    case "EMAIL":
                        if (string.IsNullOrWhiteSpace(result.Email1)) result.Email1 = value;
                        else if (string.IsNullOrWhiteSpace(result.Email2)) result.Email2 = value;
                        break;
                    case "TEL":
                        if (upperLeft.Contains("CELL") || upperLeft.Contains("MOBILE")) result.MobilePhone = value;
                        else if (upperLeft.Contains("HOME")) result.HomePhone = value;
                        else if (string.IsNullOrWhiteSpace(result.BusinessPhone)) result.BusinessPhone = value;
                        break;
                    case "ADR":
                        string[] a = value.Split(';');
                        if (a.Length > 2) result.BusinessAddressStreet = a[2];
                        if (a.Length > 3) result.BusinessAddressCity = a[3];
                        if (a.Length > 5) result.BusinessAddressPostalCode = a[5];
                        if (a.Length > 6) result.BusinessAddressCountry = a[6];
                        break;
                    case "NOTE": result.Notes = value; break;
                }
            }

            if (string.IsNullOrWhiteSpace(result.FullName))
            {
                result.FullName = ((result.FirstName ?? string.Empty) + " " + (result.LastName ?? string.Empty)).Trim();
            }
            return result;
        }

        private static List<string> UnfoldVCard(string raw)
        {
            string normalized = (raw ?? string.Empty).Replace("\r\n", "\n").Replace("\r", "\n");
            string[] source = normalized.Split('\n');
            var result = new List<string>();
            foreach (string line in source)
            {
                if ((line.StartsWith(" ") || line.StartsWith("\t")) && result.Count > 0)
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
            return !string.IsNullOrWhiteSpace(status) && status.IndexOf(" 200 ", StringComparison.Ordinal) >= 0;
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
