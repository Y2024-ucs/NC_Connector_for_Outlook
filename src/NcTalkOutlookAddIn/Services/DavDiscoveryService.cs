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
            EnsureConfiguration();
            Uri principal = DiscoverPrincipal();
            Uri addressBookHome = DiscoverAddressBookHome(principal);
            return DiscoverAddressBooks(addressBookHome);
        }

        internal IList<CalDavCalendar> DiscoverCalendars()
        {
            EnsureConfiguration();
            Uri principal = DiscoverPrincipal();
            Uri calendarHome = DiscoverCalendarHome(principal);
            return DiscoverCalendars(calendarHome);
        }

        private void EnsureConfiguration()
        {
            if (!_configuration.IsComplete())
            {
                throw new InvalidOperationException("Nextcloud configuration is incomplete.");
            }
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
                throw new InvalidOperationException("DAV principal could not be discovered.");
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

        private Uri DiscoverCalendarHome(Uri principal)
        {
            XDocument response = SendPropFind(
                principal,
                0,
                "<d:prop><cal:calendar-home-set /></d:prop>");

            string href = FirstHref(response, CalDav + "calendar-home-set");
            if (string.IsNullOrWhiteSpace(href))
            {
                throw new InvalidOperationException("CalDAV calendar-home-set could not be discovered.");
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
                XElement prop = GetSuccessfulProp(responseElement);
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

        private IList<CalDavCalendar> DiscoverCalendars(Uri home)
        {
            XDocument response = SendPropFind(
                home,
                1,
                "<d:prop>"
                + "<d:displayname />"
                + "<d:resourcetype />"
                + "<d:sync-token />"
                + "<d:getetag />"
                + "<d:current-user-privilege-set />"
                + "<cal:calendar-description />"
                + "</d:prop>");

            var result = new List<CalDavCalendar>();
            foreach (XElement responseElement in response.Descendants(Dav + "response"))
            {
                XElement prop = GetSuccessfulProp(responseElement);
                if (prop == null || !IsCalendar(prop))
                {
                    continue;
                }

                string href = (string)responseElement.Element(Dav + "href") ?? string.Empty;
                Uri resolved = ResolveDavUri(home, href);
                result.Add(new CalDavCalendar
                {
                    Href = resolved.AbsoluteUri,
                    DisplayName = ((string)prop.Element(Dav + "displayname") ?? string.Empty).Trim(),
                    Description = ((string)prop.Element(CalDav + "calendar-description") ?? string.Empty).Trim(),
                    SyncToken = ((string)prop.Element(Dav + "sync-token") ?? string.Empty).Trim(),
                    ETag = ((string)prop.Element(Dav + "getetag") ?? string.Empty).Trim(),
                    ReadOnly = IsReadOnly(prop)
                });
            }

            return result.OrderBy(c => c.DisplayName ?? string.Empty, StringComparer.CurrentCultureIgnoreCase).ToList();
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
            request.Headers["Depth"] = depth.ToString(System.Globalization.CultureInfo.InvariantCulture);
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
                        throw new InvalidOperationException("DAV server returned an empty response.");
                    }

                    return XDocument.Load(stream);
                }
            }
            catch (WebException ex)
            {
                throw new InvalidOperationException(
                    "DAV discovery request failed for " + uri.AbsoluteUri + ".",
                    ex);
            }
        }

        private static XElement GetSuccessfulProp(XElement responseElement)
        {
            if (responseElement == null)
            {
                return null;
            }
            return responseElement
                .Elements(Dav + "propstat")
                .Where(p => IsSuccessfulStatus((string)p.Element(Dav + "status")))
                .Select(p => p.Element(Dav + "prop"))
                .FirstOrDefault(p => p != null);
        }

        private static bool IsAddressBook(XElement prop)
        {
            XElement resourceType = prop.Element(Dav + "resourcetype");
            return resourceType != null && resourceType.Elements(CardDav + "addressbook").Any();
        }

        private static bool IsCalendar(XElement prop)
        {
            XElement resourceType = prop.Element(Dav + "resourcetype");
            return resourceType != null && resourceType.Elements(CalDav + "calendar").Any();
        }

        private static bool IsReadOnly(XElement prop)
        {
            XElement privilegeSet = prop.Element(Dav + "current-user-privilege-set");
            if (privilegeSet == null)
            {
                return false;
            }

            foreach (XElement privilege in privilegeSet.Descendants(Dav + "privilege"))
            {
                if (privilege.Elements(Dav + "write").Any()
                    || privilege.Elements(Dav + "write-content").Any())
                {
                    return false;
                }
            }
            return true;
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
