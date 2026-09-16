// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

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
}
