// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class CardDavCustomerSyncResult
    {
        internal int Uploaded { get; set; }
        internal int Downloaded { get; set; }
        internal int DeletedInOutlook { get; set; }
        internal int DeletedInCloud { get; set; }
        internal int Failed { get; set; }
        internal bool DeletionsBlocked { get; set; }

        internal string Summary()
        {
            string text = "Kunden: " + Uploaded + " hochgeladen, " + Downloaded + " heruntergeladen, "
                + DeletedInOutlook + " in Outlook gelöscht, " + DeletedInCloud + " in der Cloud gelöscht";
            if (Failed > 0) text += ", " + Failed + " fehlgeschlagen";
            if (DeletionsBlocked) text += ". Löschungen wurden zurückgehalten, weil zu viele Kontakte auf einmal gelöscht würden";
            return text + ".";
        }
    }

    /// <summary>
    /// Two-way sync between the Outlook contacts subfolder "Kunden" and the Nextcloud address book
    /// "IBP-Kunden". Outlook COM runs through <c>runOnUi</c> in short slices; HTTP runs on the pool.
    /// The folder itself is never deleted; contact deletions need a previously synced href.
    /// </summary>
    internal sealed class CardDavCustomerSync
    {
        internal const string FolderName = "Kunden";
        internal const string AddressBookDisplayName = "IBP-Kunden";
        internal const string AddressBookSlug = "ibp-kunden";

        private const string UidPropertyName = "NC-IBP-KUNDE-UID";
        private const string HrefPropertyName = "NC-IBP-KUNDE-HREF";
        private const string ETagPropertyName = "NC-IBP-KUNDE-ETAG";
        private const string HashPropertyName = "NC-IBP-KUNDE-HASH";
        private const string StateFileName = "carddav-customers.xml";
        private const int SliceBudgetMilliseconds = 40;

        private static readonly XNamespace Dav = "DAV:";
        private static readonly XNamespace CardDav = "urn:ietf:params:xml:ns:carddav";

        private readonly TalkServiceConfiguration _configuration;
        private readonly Outlook.Application _application;
        private readonly Func<Action, Task> _runOnUi;
        private readonly Func<Task> _yieldUi;

        internal CardDavCustomerSync(
            TalkServiceConfiguration configuration,
            Outlook.Application application,
            Func<Action, Task> runOnUi,
            Func<Task> yieldUi)
        {
            if (configuration == null) throw new ArgumentNullException("configuration");
            if (application == null) throw new ArgumentNullException("application");
            if (runOnUi == null) throw new ArgumentNullException("runOnUi");
            _configuration = configuration;
            _application = application;
            _runOnUi = runOnUi;
            _yieldUi = yieldUi ?? (() => Task.Delay(20));
        }

        /// <summary>
        /// The user's own customer book. Books shared by colleagues ("..._shared_by_...") can carry the
        /// same display name, but this sync writes and deletes, so it must never pick one of those.
        /// </summary>
        internal static bool IsCustomerAddressBook(CardDavAddressBook addressBook)
        {
            Uri uri;
            if (addressBook == null || !Uri.TryCreate(addressBook.Href ?? string.Empty, UriKind.Absolute, out uri))
            {
                return false;
            }
            string collection = uri.AbsolutePath.TrimEnd('/').Split('/').Last();
            if (string.Equals(collection, AddressBookSlug, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            return collection.IndexOf("_shared_by_", StringComparison.OrdinalIgnoreCase) < 0
                && string.Equals((addressBook.DisplayName ?? string.Empty).Trim(), AddressBookDisplayName, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// The customer book is synced by this class only; the read-only sync must not import it again.
        /// </summary>
        internal static IList<CardDavAddressBook> WithoutCustomerAddressBook(IEnumerable<CardDavAddressBook> addressBooks)
        {
            return (addressBooks ?? Enumerable.Empty<CardDavAddressBook>())
                .Where(book => !IsCustomerAddressBook(book))
                .ToList();
        }

        internal static bool IsCustomerFolderName(string folderName)
        {
            return OutlookFolderNames.Matches(folderName, FolderName);
        }

        internal async Task<CardDavCustomerSyncResult> RunAsync()
        {
            var result = new CardDavCustomerSyncResult();

            // 1. Ensure the Outlook folder and list its items (one fast table scan).
            string storeId = null;
            string folderEntryId = null;
            var entryIds = new List<string>();
            await _runOnUi(() =>
            {
                Outlook.NameSpace session = null;
                Outlook.MAPIFolder root = null;
                Outlook.MAPIFolder folder = null;
                try
                {
                    session = _application.Session;
                    root = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                    folder = CardDavReadOnlySync.EnsureSubFolder(root, FolderName);
                    if (CardDavContactFolderSync.IsManagedGroupFolder(folder))
                    {
                        throw new InvalidOperationException(
                            "Der Outlook-Ordner \"" + FolderName + "\" wird bereits als Nextcloud-Gruppenordner verwendet.");
                    }
                    storeId = folder.StoreID;
                    folderEntryId = folder.EntryID;
                    entryIds.AddRange(ListEntryIds(folder));
                }
                finally
                {
                    ComInteropScope.TryRelease(folder, LogCategories.Core, "Failed to release Outlook customer folder.");
                    ComInteropScope.TryRelease(root, LogCategories.Core, "Failed to release Outlook contacts folder after customer scan.");
                    ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session after customer scan.");
                }
            }).ConfigureAwait(false);

            // 2. Read the contacts in slices so Outlook stays responsive.
            var localEntries = new List<CardDavCustomerLocalEntry>();
            int index = 0;
            while (index < entryIds.Count)
            {
                await _runOnUi(() =>
                {
                    Outlook.NameSpace session = null;
                    try
                    {
                        session = _application.Session;
                        var stopwatch = Stopwatch.StartNew();
                        do
                        {
                            CardDavCustomerLocalEntry entry = ReadLocalEntry(session, entryIds[index], storeId);
                            if (entry != null) localEntries.Add(entry);
                            index++;
                        }
                        while (index < entryIds.Count && stopwatch.ElapsedMilliseconds < SliceBudgetMilliseconds);
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session while reading customers.");
                    }
                }).ConfigureAwait(false);
                if (index < entryIds.Count) await _yieldUi().ConfigureAwait(false);
            }

            // 3. Server side: address book, versions, plan, uploads and deletions.
            RemotePhase remote = await Task.Run(() => RunRemotePhase(localEntries)).ConfigureAwait(false);
            CardDavCustomerPlan plan = remote.Plan;
            result.DeletionsBlocked = plan.DeletionsBlocked;
            if (plan.DeletionsBlocked)
            {
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "CardDAV customer sync held back deletions (count=" + plan.BlockedDeletionHrefs.Count
                    + ", known=" + remote.KnownCount + ").");
            }

            // 4. Apply the outcome to Outlook in slices.
            List<CardDavCustomerAction> outlookActions = plan.Actions
                .Where(a => a.Kind != CardDavCustomerActionKind.DeleteRemote)
                .ToList();
            index = 0;
            while (index < outlookActions.Count)
            {
                await _runOnUi(() =>
                {
                    Outlook.NameSpace session = null;
                    try
                    {
                        session = _application.Session;
                        var stopwatch = Stopwatch.StartNew();
                        do
                        {
                            ApplyToOutlook(session, storeId, folderEntryId, outlookActions[index], remote.FinalVersions);
                            index++;
                        }
                        while (index < outlookActions.Count && stopwatch.ElapsedMilliseconds < SliceBudgetMilliseconds);
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session while applying customers.");
                    }
                }).ConfigureAwait(false);
                if (index < outlookActions.Count) await _yieldUi().ConfigureAwait(false);
            }

            foreach (CardDavCustomerAction action in plan.Actions)
            {
                if (!action.Succeeded)
                {
                    result.Failed++;
                    continue;
                }
                switch (action.Kind)
                {
                    case CardDavCustomerActionKind.UploadNew:
                    case CardDavCustomerActionKind.UploadUpdate: result.Uploaded++; break;
                    case CardDavCustomerActionKind.DownloadNew:
                    case CardDavCustomerActionKind.DownloadUpdate: result.Downloaded++; break;
                    case CardDavCustomerActionKind.DeleteLocal: result.DeletedInOutlook++; break;
                    case CardDavCustomerActionKind.DeleteRemote: result.DeletedInCloud++; break;
                }
            }

            // 5. Remember which contacts exist on both sides; only these may be deleted later.
            SaveState(remote.AddressBookHref, BuildKnownHrefs(localEntries, plan, remote.FinalVersions));
            DiagnosticsLogger.Log(
                LogCategories.Core,
                "CardDAV customer sync completed (local=" + localEntries.Count
                + ", remote=" + remote.FinalVersions.Count
                + ", uploaded=" + result.Uploaded
                + ", downloaded=" + result.Downloaded
                + ", deletedOutlook=" + result.DeletedInOutlook
                + ", deletedCloud=" + result.DeletedInCloud
                + ", failed=" + result.Failed + ").");
            return result;
        }

        private sealed class RemotePhase
        {
            internal string AddressBookHref { get; set; }
            internal CardDavCustomerPlan Plan { get; set; }
            internal Dictionary<string, string> FinalVersions { get; set; }
            internal int KnownCount { get; set; }
        }

        private sealed class RemotePayload
        {
            internal string ETag { get; set; }
            internal string Raw { get; set; }
        }

        private RemotePhase RunRemotePhase(IList<CardDavCustomerLocalEntry> localEntries)
        {
            CardDavAddressBook addressBook = new DavDiscoveryService(_configuration)
                .EnsureAddressBook(AddressBookSlug, AddressBookDisplayName, IsCustomerAddressBook);
            string bookHref = addressBook.Href.TrimEnd('/') + "/";

            HashSet<string> known = LoadState(bookHref);
            CardDavRemoteSnapshot before = LoadVersions(bookHref);
            CardDavCustomerPlan plan = CardDavCustomerPlanner.Plan(localEntries, before.Versions, known, bookHref);

            List<string> payloadHrefs = plan.Actions
                .Where(a => a.Kind == CardDavCustomerActionKind.UploadUpdate
                    || a.Kind == CardDavCustomerActionKind.DownloadNew
                    || a.Kind == CardDavCustomerActionKind.DownloadUpdate)
                .Select(a => a.Href)
                .Distinct(StringComparer.Ordinal)
                .ToList();
            Dictionary<string, RemotePayload> payloads = LoadPayloads(bookHref, payloadHrefs);

            foreach (CardDavCustomerAction action in plan.Actions)
            {
                try
                {
                    ExecuteRemote(bookHref, action, payloads);
                }
                catch (Exception ex)
                {
                    action.Succeeded = false;
                    DiagnosticsLogger.LogException(
                        LogCategories.Core,
                        "CardDAV customer action failed (kind=" + action.Kind + ").",
                        ex);
                }
            }

            CardDavRemoteSnapshot after = LoadVersions(bookHref);
            return new RemotePhase
            {
                AddressBookHref = bookHref,
                Plan = plan,
                FinalVersions = after.Versions,
                KnownCount = known.Count
            };
        }

        private void ExecuteRemote(string bookHref, CardDavCustomerAction action, IDictionary<string, RemotePayload> payloads)
        {
            RemotePayload payload;
            switch (action.Kind)
            {
                case CardDavCustomerActionKind.UploadNew:
                    HttpStatusCode created = Put(action.Href, BuildVCard(action, null), null);
                    if (created == HttpStatusCode.PreconditionFailed && !IsFreshHref(action))
                    {
                        // The old href is taken by another resource; keep both and upload under a new one.
                        action.Uid = Guid.NewGuid().ToString("D");
                        action.Href = CardDavCustomerPlanner.BuildHref(bookHref, action.Uid);
                        created = Put(action.Href, BuildVCard(action, null), null);
                    }
                    action.Succeeded = IsSuccess(created);
                    break;

                case CardDavCustomerActionKind.UploadUpdate:
                    if (!payloads.TryGetValue(action.Href, out payload)) return;
                    // A 412 means the cloud changed after the listing; the next run retries.
                    action.Succeeded = IsSuccess(Put(action.Href, BuildVCard(action, payload.Raw), payload.ETag));
                    break;

                case CardDavCustomerActionKind.DownloadNew:
                case CardDavCustomerActionKind.DownloadUpdate:
                    if (!payloads.TryGetValue(action.Href, out payload)) return;
                    string normalized = CardDavVCardSupport.Normalize(payload.Raw);
                    CardDavContactRecord record = CardDavReadOnlySync.ParseVCard(normalized);
                    record.Href = action.Href;
                    record.ETag = payload.ETag;
                    CardDavVCardSupport.CachePhoto(action.Href, normalized, _configuration);
                    action.Downloaded = record;
                    action.RemoteEtag = payload.ETag;
                    // Succeeded is set once Outlook has saved the contact.
                    break;

                case CardDavCustomerActionKind.DeleteRemote:
                    HttpStatusCode deleted = Send("DELETE", action.Href, null, action.RemoteEtag, null).StatusCode;
                    action.Succeeded = IsSuccess(deleted) || deleted == HttpStatusCode.NotFound;
                    break;

                case CardDavCustomerActionKind.DeleteLocal:
                    // Outlook side only.
                    break;
            }
        }

        private static bool IsFreshHref(CardDavCustomerAction action)
        {
            return action.Local == null
                || !string.Equals(
                    CardDavCustomerPlanner.NormalizeHref(action.Local.Href),
                    action.Href,
                    StringComparison.Ordinal);
        }

        private static string BuildVCard(CardDavCustomerAction action, string existingRaw)
        {
            return CardDavCustomerVCard.Build(
                action.Local != null ? action.Local.Fields : null,
                action.Uid,
                existingRaw,
                DateTime.UtcNow);
        }

        private HttpStatusCode Put(string href, string vcard, string ifMatch)
        {
            NcHttpResponse response = Send(
                "PUT",
                href,
                vcard,
                ifMatch,
                string.IsNullOrWhiteSpace(ifMatch) ? "*" : null,
                "text/vcard; charset=utf-8");
            return response.StatusCode;
        }

        private NcHttpResponse Send(
            string method,
            string url,
            string body,
            string ifMatch,
            string ifNoneMatch,
            string contentType = "application/xml; charset=utf-8",
            string depth = null)
        {
            var headers = new Dictionary<string, string>();
            if (!string.IsNullOrWhiteSpace(ifMatch)) headers["If-Match"] = ifMatch.Trim();
            if (!string.IsNullOrWhiteSpace(ifNoneMatch)) headers["If-None-Match"] = ifNoneMatch;
            if (!string.IsNullOrWhiteSpace(depth)) headers["Depth"] = depth;

            NcHttpResponse response = new NcHttpClient(_configuration).Send(new NcHttpRequestOptions
            {
                Method = method,
                Url = url,
                Payload = body,
                Accept = "application/xml, text/xml, text/vcard",
                ContentType = contentType,
                IncludeOcsApiHeader = false,
                ParseJson = false,
                Headers = headers
            });
            if (!response.HasHttpResponse)
            {
                throw new InvalidOperationException("CardDAV customer request failed for " + url + ".", response.TransportException);
            }
            return response;
        }

        private static bool IsSuccess(HttpStatusCode status)
        {
            int code = (int)status;
            return code >= 200 && code < 300;
        }

        private CardDavRemoteSnapshot LoadVersions(string bookHref)
        {
            string body = "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                + "<card:addressbook-query xmlns:d=\"DAV:\" xmlns:card=\"urn:ietf:params:xml:ns:carddav\">"
                + "<d:prop><d:getetag/></d:prop><card:filter/></card:addressbook-query>";
            NcHttpResponse response = Send("REPORT", bookHref, body, null, null, depth: "1");
            if ((int)response.StatusCode != 207)
            {
                throw new InvalidOperationException("CardDAV customer listing failed (" + (int)response.StatusCode + ").");
            }
            // Parse validates completeness; an incomplete listing must never drive deletions.
            return CardDavRemoteSnapshot.Parse(bookHref, XDocument.Parse(response.ResponseText ?? string.Empty));
        }

        private Dictionary<string, RemotePayload> LoadPayloads(string bookHref, IList<string> hrefs)
        {
            var result = new Dictionary<string, RemotePayload>(StringComparer.Ordinal);
            if (hrefs == null || hrefs.Count == 0) return result;

            var body = new StringBuilder();
            body.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            body.Append("<card:addressbook-multiget xmlns:d=\"DAV:\" xmlns:card=\"urn:ietf:params:xml:ns:carddav\">");
            body.Append("<d:prop><d:getetag/><card:address-data content-type=\"text/vcard\"/></d:prop>");
            foreach (string href in hrefs)
            {
                body.Append("<d:href>").Append(SecurityElementEscape(new Uri(href).AbsolutePath)).Append("</d:href>");
            }
            body.Append("</card:addressbook-multiget>");

            NcHttpResponse response = Send("REPORT", bookHref, body.ToString(), null, null, depth: "1");
            if ((int)response.StatusCode != 207)
            {
                throw new InvalidOperationException("CardDAV customer payload request failed (" + (int)response.StatusCode + ").");
            }

            XDocument document = XDocument.Parse(response.ResponseText ?? string.Empty);
            var book = new Uri(bookHref);
            foreach (XElement responseElement in document.Descendants(Dav + "response"))
            {
                XElement prop = responseElement
                    .Elements(Dav + "propstat")
                    .Where(p => ((string)p.Element(Dav + "status") ?? string.Empty).Contains(" 200 "))
                    .Select(p => p.Element(Dav + "prop"))
                    .FirstOrDefault(p => p != null);
                string href = ((string)responseElement.Element(Dav + "href") ?? string.Empty).Trim();
                Uri resolved;
                if (prop == null || href.Length == 0 || !Uri.TryCreate(book, href, out resolved)) continue;
                string raw = (string)prop.Element(CardDav + "address-data");
                if (string.IsNullOrWhiteSpace(raw)) continue;
                result[resolved.AbsoluteUri] = new RemotePayload
                {
                    ETag = ((string)prop.Element(Dav + "getetag") ?? string.Empty).Trim(),
                    Raw = raw
                };
            }
            return result;
        }

        private static string SecurityElementEscape(string value)
        {
            return System.Security.SecurityElement.Escape(value ?? string.Empty);
        }

        private static IEnumerable<string> ListEntryIds(Outlook.MAPIFolder folder)
        {
            var result = new List<string>();
            Outlook.Table table = null;
            try
            {
                table = folder.GetTable(Type.Missing, Outlook.OlTableContents.olUserItems);
                while (!table.EndOfTable)
                {
                    Outlook.Row row = null;
                    try
                    {
                        row = table.GetNextRow();
                        string entryId = Convert.ToString(row["EntryID"]);
                        if (!string.IsNullOrWhiteSpace(entryId)) result.Add(entryId);
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(row, LogCategories.Core, "Failed to release customer table row.");
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(table, LogCategories.Core, "Failed to release customer folder table.");
            }
            return result;
        }

        private static CardDavCustomerLocalEntry ReadLocalEntry(Outlook.NameSpace session, string entryId, string storeId)
        {
            Outlook.ContactItem contact = CardDavReadOnlySync.TryGetContactById(session, entryId, storeId);
            if (contact == null) return null;
            try
            {
                CardDavContactRecord fields = ReadFields(contact);
                return new CardDavCustomerLocalEntry
                {
                    EntryId = entryId,
                    Uid = CardDavReadOnlySync.ReadUserProperty(contact, UidPropertyName),
                    Href = CardDavReadOnlySync.ReadUserProperty(contact, HrefPropertyName),
                    StoredEtag = CardDavReadOnlySync.ReadUserProperty(contact, ETagPropertyName),
                    StoredHash = CardDavReadOnlySync.ReadUserProperty(contact, HashPropertyName),
                    CurrentHash = CardDavCustomerVCard.ComputeHash(fields),
                    Fields = fields
                };
            }
            finally
            {
                ComInteropScope.TryRelease(contact, LogCategories.Core, "Failed to release customer contact after read.");
            }
        }

        private static CardDavContactRecord ReadFields(Outlook.ContactItem contact)
        {
            return new CardDavContactRecord
            {
                FullName = contact.FullName,
                FirstName = contact.FirstName,
                LastName = contact.LastName,
                Company = contact.CompanyName,
                JobTitle = contact.JobTitle,
                Email1 = SmtpOrEmpty(contact.Email1Address, contact.Email1AddressType),
                Email2 = SmtpOrEmpty(contact.Email2Address, contact.Email2AddressType),
                BusinessPhone = contact.BusinessTelephoneNumber,
                BusinessPhone2 = contact.Business2TelephoneNumber,
                MobilePhone = contact.MobileTelephoneNumber,
                HomePhone = contact.HomeTelephoneNumber,
                HomePhone2 = contact.Home2TelephoneNumber,
                BusinessFax = contact.BusinessFaxNumber,
                HomeFax = contact.HomeFaxNumber,
                OtherPhone = contact.OtherTelephoneNumber,
                PagerPhone = contact.PagerNumber,
                CompanyMainPhone = contact.CompanyMainTelephoneNumber,
                CarPhone = contact.CarTelephoneNumber,
                BusinessAddressStreet = contact.BusinessAddressStreet,
                BusinessAddressCity = contact.BusinessAddressCity,
                BusinessAddressState = contact.BusinessAddressState,
                BusinessAddressPostalCode = contact.BusinessAddressPostalCode,
                BusinessAddressCountry = contact.BusinessAddressCountry,
                Notes = contact.Body
            };
        }

        private static string SmtpOrEmpty(string address, string addressType)
        {
            // Exchange-internal addresses (/o=...) mean nothing to other CardDAV clients.
            return string.Equals(addressType, "EX", StringComparison.OrdinalIgnoreCase)
                ? string.Empty
                : (address ?? string.Empty);
        }

        private void ApplyToOutlook(
            Outlook.NameSpace session,
            string storeId,
            string folderEntryId,
            CardDavCustomerAction action,
            IDictionary<string, string> finalVersions)
        {
            try
            {
                switch (action.Kind)
                {
                    case CardDavCustomerActionKind.UploadNew:
                    case CardDavCustomerActionKind.UploadUpdate:
                        if (action.Succeeded) WriteUploadMarkers(session, storeId, action, finalVersions);
                        break;
                    case CardDavCustomerActionKind.DownloadNew:
                    case CardDavCustomerActionKind.DownloadUpdate:
                        if (action.Downloaded != null) action.Succeeded = WriteDownload(session, storeId, folderEntryId, action);
                        break;
                    case CardDavCustomerActionKind.DeleteLocal:
                        action.Succeeded = DeleteLocal(session, storeId, action);
                        break;
                }
            }
            catch (Exception ex)
            {
                action.Succeeded = false;
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "CardDAV customer Outlook update failed (kind=" + action.Kind + ").",
                    ex);
            }
        }

        private static void WriteUploadMarkers(
            Outlook.NameSpace session,
            string storeId,
            CardDavCustomerAction action,
            IDictionary<string, string> finalVersions)
        {
            Outlook.ContactItem contact = CardDavReadOnlySync.TryGetContactById(session, action.Local.EntryId, storeId);
            if (contact == null) return;
            try
            {
                string etag;
                finalVersions.TryGetValue(action.Href, out etag);
                CardDavReadOnlySync.WriteUserProperty(contact, UidPropertyName, action.Uid);
                CardDavReadOnlySync.WriteUserProperty(contact, HrefPropertyName, action.Href);
                CardDavReadOnlySync.WriteUserProperty(contact, ETagPropertyName, etag ?? string.Empty);
                // The hash of what was read, not of the current state: edits made meanwhile upload next run.
                CardDavReadOnlySync.WriteUserProperty(contact, HashPropertyName, action.Local.CurrentHash);
                contact.Save();
            }
            finally
            {
                ComInteropScope.TryRelease(contact, LogCategories.Core, "Failed to release customer contact after upload.");
            }
        }

        private static bool WriteDownload(
            Outlook.NameSpace session,
            string storeId,
            string folderEntryId,
            CardDavCustomerAction action)
        {
            Outlook.ContactItem contact = null;
            Outlook.MAPIFolder folder = null;
            Outlook.Items items = null;
            try
            {
                if (action.Kind == CardDavCustomerActionKind.DownloadUpdate)
                {
                    contact = CardDavReadOnlySync.TryGetContactById(session, action.Local.EntryId, storeId);
                    // Deleted or edited in Outlook since the scan: leave it, the next run decides again.
                    if (contact == null
                        || !string.Equals(CardDavCustomerVCard.ComputeHash(ReadFields(contact)), action.Local.CurrentHash, StringComparison.Ordinal))
                    {
                        return false;
                    }
                }
                else
                {
                    folder = session.GetFolderFromID(folderEntryId, storeId);
                    items = folder.Items;
                    contact = items.Add(Outlook.OlItemType.olContactItem) as Outlook.ContactItem;
                    if (contact == null) return false;
                }

                CardDavReadOnlySync.ApplyContactFields(contact, action.Downloaded);
                CardDavVCardSupport.ApplyPhoto(contact, action.Href);
                CardDavReadOnlySync.WriteUserProperty(contact, UidPropertyName, action.Downloaded.Uid);
                CardDavReadOnlySync.WriteUserProperty(contact, HrefPropertyName, action.Href);
                CardDavReadOnlySync.WriteUserProperty(contact, ETagPropertyName, action.RemoteEtag);
                CardDavReadOnlySync.WriteUserProperty(contact, HashPropertyName, CardDavCustomerVCard.ComputeHash(ReadFields(contact)));
                contact.Save();
                return true;
            }
            finally
            {
                ComInteropScope.TryRelease(contact, LogCategories.Core, "Failed to release downloaded customer contact.");
                ComInteropScope.TryRelease(items, LogCategories.Core, "Failed to release customer folder items.");
                ComInteropScope.TryRelease(folder, LogCategories.Core, "Failed to release customer folder after download.");
            }
        }

        private static bool DeleteLocal(Outlook.NameSpace session, string storeId, CardDavCustomerAction action)
        {
            Outlook.ContactItem contact = CardDavReadOnlySync.TryGetContactById(session, action.Local.EntryId, storeId);
            if (contact == null) return true;
            try
            {
                // Edited since the scan: keep it, the next run uploads it again (Outlook wins).
                if (!string.Equals(CardDavCustomerVCard.ComputeHash(ReadFields(contact)), action.Local.CurrentHash, StringComparison.Ordinal))
                {
                    return false;
                }
                // Delete() moves the contact to "Deleted Items", so the user can restore it.
                contact.Delete();
                return true;
            }
            finally
            {
                ComInteropScope.TryRelease(contact, LogCategories.Core, "Failed to release deleted customer contact.");
            }
        }

        private static HashSet<string> BuildKnownHrefs(
            IEnumerable<CardDavCustomerLocalEntry> localEntries,
            CardDavCustomerPlan plan,
            IDictionary<string, string> finalVersions)
        {
            var final = new HashSet<string>(finalVersions.Keys, StringComparer.Ordinal);
            var inOutlook = new HashSet<string>(StringComparer.Ordinal);
            var actionsByEntry = plan.Actions
                .Where(a => a.Local != null)
                .ToDictionary(a => a.Local, a => a);

            foreach (CardDavCustomerLocalEntry entry in localEntries)
            {
                CardDavCustomerAction action;
                actionsByEntry.TryGetValue(entry, out action);
                if (action == null)
                {
                    inOutlook.Add(CardDavCustomerPlanner.NormalizeHref(entry.Href));
                }
                else if (action.Kind == CardDavCustomerActionKind.DeleteLocal)
                {
                    if (!action.Succeeded) inOutlook.Add(action.Href);
                }
                else if (action.Succeeded
                    || action.Kind == CardDavCustomerActionKind.UploadUpdate
                    || action.Kind == CardDavCustomerActionKind.DownloadUpdate)
                {
                    inOutlook.Add(action.Kind == CardDavCustomerActionKind.UploadNew
                        ? action.Href
                        : CardDavCustomerPlanner.NormalizeHref(entry.Href));
                }
            }
            foreach (CardDavCustomerAction action in plan.Actions)
            {
                if (action.Kind == CardDavCustomerActionKind.DownloadNew && action.Succeeded) inOutlook.Add(action.Href);
            }

            var known = new HashSet<string>(inOutlook.Where(final.Contains), StringComparer.Ordinal);
            // Deletions that were held back or failed must still count as synced, or they would come back.
            foreach (string href in plan.BlockedDeletionHrefs) known.Add(href);
            foreach (CardDavCustomerAction action in plan.Actions)
            {
                if ((action.Kind == CardDavCustomerActionKind.DeleteRemote || action.Kind == CardDavCustomerActionKind.DeleteLocal)
                    && !action.Succeeded)
                {
                    known.Add(action.Href);
                }
            }
            return known;
        }

        private static string GetStatePath()
        {
            return Path.Combine(AppDataPaths.EnsureLocalRootDirectory(), StateFileName);
        }

        private static HashSet<string> LoadState(string bookHref)
        {
            var result = new HashSet<string>(StringComparer.Ordinal);
            string path = GetStatePath();
            if (!File.Exists(path)) return result;
            try
            {
                XDocument document = XDocument.Load(path);
                XElement root = document.Root;
                // A recreated or different address book starts without deletion evidence.
                if (root == null || !string.Equals((string)root.Attribute("addressBook"), bookHref, StringComparison.Ordinal))
                {
                    return result;
                }
                foreach (XElement contact in root.Elements("Contact"))
                {
                    string href = CardDavCustomerPlanner.NormalizeHref((string)contact.Attribute("href"));
                    if (href.Length > 0) result.Add(href);
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to load CardDAV customer sync state.", ex);
            }
            return result;
        }

        private static void SaveState(string bookHref, IEnumerable<string> knownHrefs)
        {
            var root = new XElement("CardDavCustomers", new XAttribute("addressBook", bookHref ?? string.Empty));
            foreach (string href in knownHrefs.OrderBy(h => h, StringComparer.Ordinal))
            {
                root.Add(new XElement("Contact", new XAttribute("href", href)));
            }
            string path = GetStatePath();
            string temp = path + ".tmp";
            var settings = new XmlWriterSettings { Indent = true, Encoding = new UTF8Encoding(false) };
            using (XmlWriter writer = XmlWriter.Create(temp, settings))
            {
                new XDocument(root).Save(writer);
            }
            if (File.Exists(path)) File.Replace(temp, path, null);
            else File.Move(temp, path);
        }
    }
}
