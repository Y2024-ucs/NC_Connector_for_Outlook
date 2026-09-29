// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Xml.Linq;
using NcTalkOutlookAddIn.Models;

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class DavDiscoveryService
    {
        private static readonly XNamespace Dav = "DAV:";
        private static readonly XNamespace CardDav = "urn:ietf:params:xml:ns:carddav";
        private static readonly XNamespace CalDav = "urn:ietf:params:xml:ns:caldav";

        private readonly TalkServiceConfiguration _configuration;

        internal DavDiscoveryService(TalkServiceConfiguration configuration)
        {
            if (configuration == null)
            {
                throw new ArgumentNullException("configuration");
            }
            _configuration = configuration;
        }

        internal IList<CardDavAddressBook> DiscoverAddressBooks()
        {
            if (!_configuration.IsComplete())
            {
                throw new InvalidOperationException("Nextcloud configuration is incomplete.");
            }

            Uri principal = DiscoverPrincipal();
            Uri addressBookHome = DiscoverAddressBookHome(principal);
            return DiscoverAddressBooks(addressBookHome);
        }

        /// <summary>
        /// Returns the user's address book matching <paramref name="match"/>, creating it via MKCOL when missing.
        /// </summary>
        internal CardDavAddressBook EnsureAddressBook(
            string slug,
            string displayName,
            Func<CardDavAddressBook, bool> match)
        {
            if (!_configuration.IsComplete())
            {
                throw new InvalidOperationException("Nextcloud configuration is incomplete.");
            }

            Uri addressBookHome = DiscoverAddressBookHome(DiscoverPrincipal());
            CardDavAddressBook existing = DiscoverAddressBooks(addressBookHome).FirstOrDefault(match);
            if (existing != null)
            {
                return existing;
            }

            Uri target = new Uri(new Uri(addressBookHome.AbsoluteUri.TrimEnd('/') + "/"), slug + "/");
            string body = "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                + "<d:mkcol xmlns:d=\"DAV:\" xmlns:card=\"urn:ietf:params:xml:ns:carddav\"><d:set><d:prop>"
                + "<d:resourcetype><d:collection/><card:addressbook/></d:resourcetype>"
                + "<d:displayname>" + System.Security.SecurityElement.Escape(displayName) + "</d:displayname>"
                + "</d:prop></d:set></d:mkcol>";
            NcHttpResponse response = new NcHttpClient(_configuration).Send(new NcHttpRequestOptions
            {
                Method = "MKCOL",
                Url = target.AbsoluteUri,
                Payload = body,
                Accept = "application/xml, text/xml",
                ContentType = "application/xml; charset=utf-8",
                IncludeOcsApiHeader = false,
                ParseJson = false
            });
            // 405: the collection exists already (e.g. renamed by the user); find it again below.
            if (!response.HasHttpResponse
                || ((int)response.StatusCode != 201 && response.StatusCode != HttpStatusCode.MethodNotAllowed))
            {
                throw new InvalidOperationException(
                    "CardDAV address book '" + displayName + "' could not be created ("
                    + (response.HasHttpResponse ? ((int)response.StatusCode).ToString() : "no response") + ").");
            }
            Utilities.DiagnosticsLogger.Log(
                Utilities.LogCategories.Core,
                "CardDAV address book ensured (name=" + displayName + ", status=" + (int)response.StatusCode + ").");

            CardDavAddressBook created = DiscoverAddressBooks(addressBookHome).FirstOrDefault(
                book => match(book) || string.Equals(
                    book.Href.TrimEnd('/'),
                    target.AbsoluteUri.TrimEnd('/'),
                    StringComparison.OrdinalIgnoreCase));
            if (created == null)
            {
                throw new InvalidOperationException("CardDAV address book '" + displayName + "' was not found after creation.");
            }
            return created;
        }

        /// <summary>
        /// URL of the user's calendar with the given collection name (Nextcloud's default is "personal"),
        /// found via calendar-home-set. Returns null when that calendar does not exist.
        /// </summary>
        internal string DiscoverCalendarUrl(string collectionName)
        {
            if (!_configuration.IsComplete())
            {
                throw new InvalidOperationException("Nextcloud configuration is incomplete.");
            }

            Uri principal = DiscoverPrincipal();
            XDocument homeResponse = SendPropFind(principal, 0, "<d:prop><cal:calendar-home-set /></d:prop>");
            string homeHref = FirstHref(homeResponse, CalDav + "calendar-home-set");
            if (string.IsNullOrWhiteSpace(homeHref))
            {
                throw new InvalidOperationException("CalDAV calendar-home-set could not be discovered.");
            }

            Uri home = ResolveDavUri(principal, homeHref);
            Uri calendar = new Uri(new Uri(home.AbsoluteUri.TrimEnd('/') + "/"), collectionName.Trim('/') + "/");
            XDocument calendarResponse;
            try
            {
                calendarResponse = SendPropFind(calendar, 0, "<d:prop><d:resourcetype /></d:prop>");
            }
            catch (InvalidOperationException ex)
            {
                var web = ex.InnerException as WebException;
                var http = web != null ? web.Response as HttpWebResponse : null;
                if (http != null && http.StatusCode == HttpStatusCode.NotFound)
                {
                    return null;
                }
                throw;
            }
            bool isCalendar = calendarResponse
                .Descendants(Dav + "resourcetype")
                .Any(type => type.Elements(CalDav + "calendar").Any());
            return isCalendar ? calendar.AbsoluteUri : null;
        }

        private Uri DiscoverPrincipal()
        {
            // This is a Nextcloud-specific add-in, so use the canonical DAV root
            // directly instead of relying on a .well-known redirect that may drop
            // the Authorization header in older .NET Framework stacks.
            Uri davRoot = BuildNextcloudDavRoot();
            XDocument response = SendPropFind(
                davRoot,
                0,
                "<d:prop><d:current-user-principal /></d:prop>");

            string href = FirstHref(response, Dav + "current-user-principal");
            if (string.IsNullOrWhiteSpace(href))
            {
                throw new InvalidOperationException("CardDAV principal could not be discovered.");
            }

            return ResolveDavUri(davRoot, href);
        }

        private Uri DiscoverAddressBookHome(Uri principal)
        {
            XDocument response = SendPropFind(
                principal,
                0,
                "<d:prop><card:addressbook-home-set /></d:prop>");

            string href = FirstHref(response, CardDav + "addressbook-home-set");
            if (string.IsNullOrWhiteSpace(href))
            {
                throw new InvalidOperationException("CardDAV addressbook-home-set could not be discovered.");
            }

            return ResolveDavUri(principal, href);
        }

        private IList<CardDavAddressBook> DiscoverAddressBooks(Uri home)
        {
            XDocument response = SendPropFind(
                home,
                1,
                "<d:prop>"
                + "<d:displayname />"
                + "<d:resourcetype />"
                + "<d:sync-token />"
                + "<d:getetag />"
                + "<card:addressbook-description />"
                + "</d:prop>");

            var result = new List<CardDavAddressBook>();
            foreach (XElement responseElement in response.Descendants(Dav + "response"))
            {
                XElement prop = responseElement
                    .Elements(Dav + "propstat")
                    .Where(p => IsSuccessfulStatus((string)p.Element(Dav + "status")))
                    .Select(p => p.Element(Dav + "prop"))
                    .FirstOrDefault(p => p != null);

                if (prop == null || !IsAddressBook(prop))
                {
                    continue;
                }

                string href = (string)responseElement.Element(Dav + "href") ?? string.Empty;
                Uri resolved = ResolveDavUri(home, href);

                result.Add(new CardDavAddressBook
                {
                    Href = resolved.AbsoluteUri,
                    DisplayName = ((string)prop.Element(Dav + "displayname") ?? string.Empty).Trim(),
                    Description = ((string)prop.Element(CardDav + "addressbook-description") ?? string.Empty).Trim(),
                    SyncToken = ((string)prop.Element(Dav + "sync-token") ?? string.Empty).Trim(),
                    ETag = ((string)prop.Element(Dav + "getetag") ?? string.Empty).Trim()
                });
            }

            return result;
        }

        private XDocument SendPropFind(Uri uri, int depth, string propertyXml)
        {
            string body = "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                          + "<d:propfind xmlns:d=\"DAV:\" xmlns:card=\"urn:ietf:params:xml:ns:carddav\" xmlns:cal=\"urn:ietf:params:xml:ns:caldav\">"
                          + propertyXml
                          + "</d:propfind>";

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uri);
            request.Method = "PROPFIND";
            request.ContentType = "application/xml; charset=utf-8";
            request.Accept = "application/xml, text/xml";
            request.Headers["Depth"] = depth.ToString();
            request.Headers[HttpRequestHeader.Authorization] = Utilities.HttpAuthUtilities.BuildBasicAuthHeader(
                _configuration.Username,
                _configuration.AppPassword);

            byte[] payload = Encoding.UTF8.GetBytes(body);
            request.ContentLength = payload.Length;
            using (Stream requestStream = request.GetRequestStream())
            {
                requestStream.Write(payload, 0, payload.Length);
            }

            try
            {
                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                using (Stream stream = response.GetResponseStream())
                {
                    if (stream == null)
                    {
                        throw new InvalidOperationException("CardDAV server returned an empty response.");
                    }

                    return XDocument.Load(stream);
                }
            }
            catch (WebException ex)
            {
                throw new InvalidOperationException(
                    "CardDAV discovery request failed for " + uri.AbsoluteUri + ".",
                    ex);
            }
        }

        private static bool IsAddressBook(XElement prop)
        {
            XElement resourceType = prop.Element(Dav + "resourcetype");
            return resourceType != null && resourceType.Elements(CardDav + "addressbook").Any();
        }

        private static bool IsSuccessfulStatus(string status)
        {
            return !string.IsNullOrWhiteSpace(status)
                && status.IndexOf(" 200 ", StringComparison.Ordinal) >= 0;
        }

        private static string FirstHref(XDocument document, XName containerName)
        {
            return document
                .Descendants(containerName)
                .Select(e => (string)e.Element(Dav + "href"))
                .FirstOrDefault(v => !string.IsNullOrWhiteSpace(v));
        }

        private Uri BuildNextcloudDavRoot()
        {
            string baseUrl = _configuration.GetNormalizedBaseUrl().TrimEnd('/');
            return new Uri(baseUrl + "/remote.php/dav/", UriKind.Absolute);
        }

        private static Uri ResolveDavUri(Uri baseUri, string href)
        {
            if (string.IsNullOrWhiteSpace(href))
            {
                return baseUri;
            }

            Uri absolute;
            return Uri.TryCreate(href, UriKind.Absolute, out absolute)
                ? absolute
                : new Uri(baseUri, href);
        }
    }
}
