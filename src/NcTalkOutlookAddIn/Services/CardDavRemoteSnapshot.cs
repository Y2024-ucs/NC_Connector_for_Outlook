// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Linq;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

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

    internal static class CardDavContactFolderSync
    {
        private const string HrefPropertyName = "NC-CardDAV-HREF";
        private const string ETagPropertyName = "NC-CardDAV-ETAG";
        private const string GroupCopyPropertyName = "NC-CardDAV-GROUP-COPY";
        private const string GroupCopyNamePropertyName = "NC-CardDAV-GROUP-COPY-NAME";
        private const string ManagedDistributionListPropertyName = "NC-CardDAV-GROUP";
        private const string PublicStringsBase = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/";
        private const string ManagedFolderProperty = PublicStringsBase + "NC-CardDAV-GROUP-FOLDER";
        private const string ManagedFolderNameProperty = PublicStringsBase + "NC-CardDAV-GROUP-FOLDER-NAME";
        private const string ManagedFolderSignatureProperty = PublicStringsBase + "NC-CardDAV-GROUP-FOLDER-SIGNATURE";
        private static readonly XNamespace Dav = "DAV:";
        private static readonly XNamespace CardDav = "urn:ietf:params:xml:ns:carddav";

        private sealed class SourceContactReference
        {
            internal string EntryId { get; set; }
            internal string StoreId { get; set; }
            internal string ETag { get; set; }
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
            var activeGroups = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (List<string> categories in categoriesByHref.Values)
            {
                if (categories == null) continue;
                foreach (string category in categories)
                {
                    string name = (category ?? string.Empty).Trim();
                    if (!string.IsNullOrWhiteSpace(name)) activeGroups.Add(name);
                }
            }

            Outlook.NameSpace session = null;
            Outlook.MAPIFolder defaultContacts = null;
            try
            {
                session = outlookApplication.Session;
                defaultContacts = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                Dictionary<string, SourceContactReference> sourceContacts = LoadSourceContacts(defaultContacts);
                GroupFolderState state = GroupFolderState.Load(defaultContacts.EntryID);
                List<SubfolderInfo> subfolders = ScanSubfolders(defaultContacts);
                DeleteObsoleteManagedFolders(session, defaultContacts.StoreID, subfolders, activeGroups, state);

                int folderCount = 0;
                foreach (string groupName in activeGroups.OrderBy(name => name, StringComparer.CurrentCultureIgnoreCase))
                {
                    Outlook.MAPIFolder groupFolder = null;
                    try
                    {
                        groupFolder = EnsureManagedGroupFolder(session, defaultContacts, groupName, subfolders, state);
                        if (groupFolder == null) continue;

                        DeleteManagedDistributionLists(groupFolder);

                        string signature = BuildGroupFolderSignature(groupName, sourceContacts, categoriesByHref);
                        string previousSignature = ReadFolderProperty(groupFolder, ManagedFolderSignatureProperty);
                        if (!string.IsNullOrWhiteSpace(signature)
                            && string.Equals(signature, previousSignature, StringComparison.Ordinal))
                        {
                            folderCount++;
                            DiagnosticsLogger.Log(
                                LogCategories.Core,
                                "CardDAV contact group folder unchanged (group=" + groupName + ").");
                            continue;
                        }

                        ClearManagedCopies(groupFolder);
                        int expected = CountGroupMembers(groupName, categoriesByHref);
                        int copied = CopyGroupMembers(session, groupFolder, groupName, sourceContacts, categoriesByHref);
                        if (copied == expected)
                        {
                            WriteFolderProperty(groupFolder, ManagedFolderSignatureProperty, signature);
                        }
                        else
                        {
                            WriteFolderProperty(groupFolder, ManagedFolderSignatureProperty, string.Empty);
                        }

                        folderCount++;
                        DiagnosticsLogger.Log(
                            LogCategories.Core,
                            "CardDAV contact group folder synchronized (group=" + groupName
                            + ", contacts=" + copied + ", expected=" + expected + ").");
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(groupFolder, LogCategories.Core, "Failed to release CardDAV group folder.");
                    }
                }

                state.Save();
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "CardDAV contact group folder reconciliation completed (folders=" + folderCount + ").");
                return folderCount;
            }
            finally
            {
                ComInteropScope.TryRelease(defaultContacts, LogCategories.Core, "Failed to release default contacts folder after CardDAV group folder sync.");
                ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session after CardDAV group folder sync.");
            }
        }

        private static Dictionary<string, SourceContactReference> LoadSourceContacts(Outlook.MAPIFolder defaultContacts)
        {
            var result = new Dictionary<string, SourceContactReference>(StringComparer.OrdinalIgnoreCase);
            CollectSourceContacts(defaultContacts, result);

            Outlook.Folders folders = null;
            try
            {
                folders = defaultContacts.Folders;
                int count = folders != null ? folders.Count : 0;
                for (int index = 1; index <= count; index++)
                {
                    Outlook.MAPIFolder folder = null;
                    try
                    {
                        folder = folders[index];
                        if (folder != null && !IsManagedGroupFolder(folder))
                        {
                            CollectSourceContacts(folder, result);
                        }
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(folder, LogCategories.Core, "Failed to release source contacts subfolder during CardDAV group folder scan.");
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(folders, LogCategories.Core, "Failed to release source contact folders during CardDAV group folder scan.");
            }
            return result;
        }

        private static void CollectSourceContacts(
            Outlook.MAPIFolder folder,
            IDictionary<string, SourceContactReference> result)
        {
            Outlook.Items items = null;
            try
            {
                items = folder.Items;
                string storeId = folder.StoreID;
                int count = items != null ? items.Count : 0;
                for (int index = 1; index <= count; index++)
                {
                    object raw = null;
                    Outlook.ContactItem contact = null;
                    try
                    {
                        raw = items[index];
                        contact = raw as Outlook.ContactItem;
                        if (contact == null) continue;

                        string href = ReadContactProperty(contact, HrefPropertyName);
                        if (string.IsNullOrWhiteSpace(href)
                            || string.IsNullOrWhiteSpace(contact.EntryID)
                            || result.ContainsKey(href.Trim()))
                        {
                            continue;
                        }
                        result[href.Trim()] = new SourceContactReference
                        {
                            EntryId = contact.EntryID,
                            StoreId = storeId,
                            ETag = ReadContactProperty(contact, ETagPropertyName)
                        };
                    }
                    finally
                    {
                        if (contact != null)
                        {
                            ComInteropScope.TryRelease(contact, LogCategories.Core, "Failed to release source contact during CardDAV group folder scan.");
                        }
                        else
                        {
                            ComInteropScope.TryRelease(raw, LogCategories.Core, "Failed to release non-contact during CardDAV group folder scan.");
                        }
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(items, LogCategories.Core, "Failed to release source items during CardDAV group folder scan.");
            }
        }

        private static int CopyGroupMembers(
            Outlook.NameSpace session,
            Outlook.MAPIFolder targetFolder,
            string groupName,
            IDictionary<string, SourceContactReference> sourceContacts,
            IDictionary<string, List<string>> categoriesByHref)
        {
            int copiedCount = 0;
            foreach (KeyValuePair<string, List<string>> pair in categoriesByHref)
            {
                if (pair.Value == null || !pair.Value.Any(category => string.Equals(
                    (category ?? string.Empty).Trim(), groupName, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                SourceContactReference sourceReference;
                if (!sourceContacts.TryGetValue(pair.Key, out sourceReference)
                    || sourceReference == null
                    || string.IsNullOrWhiteSpace(sourceReference.EntryId))
                {
                    DiagnosticsLogger.Log(
                        LogCategories.Core,
                        "CardDAV group folder source contact missing (group=" + groupName + ", href=" + pair.Key + ").");
                    continue;
                }

                Outlook.ContactItem source = null;
                Outlook.ContactItem copy = null;
                object moved = null;
                try
                {
                    source = session.GetItemFromID(sourceReference.EntryId, sourceReference.StoreId) as Outlook.ContactItem;
                    if (source == null) continue;
                    copy = source.Copy() as Outlook.ContactItem;
                    if (copy == null) continue;

                    WriteContactProperty(copy, GroupCopyPropertyName, "1");
                    WriteContactProperty(copy, GroupCopyNamePropertyName, groupName);
                    copy.Save();
                    moved = copy.Move(targetFolder);
                    copiedCount++;
                }
                finally
                {
                    ComInteropScope.TryRelease(moved, LogCategories.Core, "Failed to release moved CardDAV group folder contact.");
                    ComInteropScope.TryRelease(copy, LogCategories.Core, "Failed to release copied CardDAV group folder contact.");
                    ComInteropScope.TryRelease(source, LogCategories.Core, "Failed to release source CardDAV group folder contact.");
                }
            }
            return copiedCount;
        }

        private static int CountGroupMembers(
            string groupName,
            IDictionary<string, List<string>> categoriesByHref)
        {
            int count = 0;
            foreach (KeyValuePair<string, List<string>> pair in categoriesByHref)
            {
                if (pair.Value != null && pair.Value.Any(category => string.Equals(
                    (category ?? string.Empty).Trim(), groupName, StringComparison.OrdinalIgnoreCase)))
                {
                    count++;
                }
            }
            return count;
        }

        private static string BuildGroupFolderSignature(
            string groupName,
            IDictionary<string, SourceContactReference> sourceContacts,
            IDictionary<string, List<string>> categoriesByHref)
        {
            var parts = new List<string>();
            foreach (KeyValuePair<string, List<string>> pair in categoriesByHref)
            {
                if (pair.Value == null || !pair.Value.Any(category => string.Equals(
                    (category ?? string.Empty).Trim(), groupName, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                SourceContactReference sourceReference;
                string etag = sourceContacts.TryGetValue(pair.Key, out sourceReference) && sourceReference != null
                    ? sourceReference.ETag ?? string.Empty
                    : string.Empty;
                parts.Add((pair.Key ?? string.Empty).Trim().ToLowerInvariant() + "|" + etag.Trim());
            }
            parts.Sort(StringComparer.Ordinal);
            return StableHash(groupName + "\n" + string.Join("\n", parts.ToArray()));
        }

        private static string StableHash(string value)
        {
            unchecked
            {
                ulong hash = 14695981039346656037UL;
                byte[] bytes = Encoding.UTF8.GetBytes(value ?? string.Empty);
                for (int index = 0; index < bytes.Length; index++)
                {
                    hash ^= bytes[index];
                    hash *= 1099511628211UL;
                }
                return hash.ToString("X16");
            }
        }

        private static void ClearManagedCopies(Outlook.MAPIFolder folder)
        {
            Outlook.Items items = null;
            try
            {
                items = folder.Items;
                for (int index = items.Count; index >= 1; index--)
                {
                    object raw = null;
                    Outlook.ContactItem contact = null;
                    try
                    {
                        raw = items[index];
                        contact = raw as Outlook.ContactItem;
                        if (contact != null
                            && string.Equals(ReadContactProperty(contact, GroupCopyPropertyName), "1", StringComparison.Ordinal))
                        {
                            contact.Delete();
                        }
                    }
                    finally
                    {
                        if (contact != null)
                        {
                            ComInteropScope.TryRelease(contact, LogCategories.Core, "Failed to release managed CardDAV group folder contact during cleanup.");
                        }
                        else
                        {
                            ComInteropScope.TryRelease(raw, LogCategories.Core, "Failed to release non-contact during CardDAV group folder cleanup.");
                        }
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(items, LogCategories.Core, "Failed to release CardDAV group folder items during cleanup.");
            }
        }

        private static void DeleteManagedDistributionLists(Outlook.MAPIFolder folder)
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
                            && string.Equals(ReadDistributionListProperty(list, ManagedDistributionListPropertyName), "1", StringComparison.Ordinal))
                        {
                            list.Delete();
                        }
                    }
                    finally
                    {
                        if (list != null)
                        {
                            ComInteropScope.TryRelease(list, LogCategories.Core, "Failed to release stale CardDAV group distribution list.");
                        }
                        else
                        {
                            ComInteropScope.TryRelease(raw, LogCategories.Core, "Failed to release non-distribution-list during stale CardDAV group cleanup.");
                        }
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(items, LogCategories.Core, "Failed to release CardDAV group folder items during distribution-list cleanup.");
            }
        }

        private enum FolderMark
        {
            Unmanaged,
            Managed,
            // Outlook sometimes fails to return folder properties; such a folder is neither ours nor foreign.
            Unknown
        }

        private sealed class SubfolderInfo
        {
            internal string EntryId { get; set; }
            internal string Name { get; set; }
            internal FolderMark Mark { get; set; }
            internal string GroupName { get; set; }
        }

        private static List<SubfolderInfo> ScanSubfolders(Outlook.MAPIFolder parent)
        {
            var result = new List<SubfolderInfo>();
            Outlook.Folders folders = null;
            try
            {
                folders = parent.Folders;
                int count = folders != null ? folders.Count : 0;
                for (int index = 1; index <= count; index++)
                {
                    Outlook.MAPIFolder folder = null;
                    try
                    {
                        folder = folders[index];
                        if (folder == null) continue;
                        string groupName;
                        FolderMark mark = ReadFolderMark(folder, out groupName);
                        result.Add(new SubfolderInfo
                        {
                            EntryId = folder.EntryID,
                            Name = folder.Name ?? string.Empty,
                            Mark = mark,
                            GroupName = groupName
                        });
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(folder, LogCategories.Core, "Failed to release contact folder while scanning CardDAV group folders.");
                    }
                }
            }
            finally
            {
                ComInteropScope.TryRelease(folders, LogCategories.Core, "Failed to release folders collection after CardDAV group folder scan.");
            }
            return result;
        }

        /// <summary>
        /// Finds the group's folder without trusting a single property read: the remembered EntryID wins,
        /// then the marker, then a same-named folder whose marker could not be read. Creates one only
        /// when none of these exist, and removes duplicates left by earlier failed lookups.
        /// </summary>
        private static Outlook.MAPIFolder EnsureManagedGroupFolder(
            Outlook.NameSpace session,
            Outlook.MAPIFolder parent,
            string groupName,
            List<SubfolderInfo> subfolders,
            GroupFolderState state)
        {
            string storeId = parent.StoreID;
            string rememberedId = state.Get(groupName);
            List<SubfolderInfo> marked = subfolders
                .Where(f => f.Mark == FolderMark.Managed
                    && string.Equals(f.GroupName, groupName, StringComparison.OrdinalIgnoreCase))
                .ToList();

            SubfolderInfo chosen = subfolders.FirstOrDefault(f => rememberedId.Length > 0
                && string.Equals(f.EntryId, rememberedId, StringComparison.Ordinal));
            bool owned = chosen != null || marked.Count > 0;
            if (chosen == null) chosen = marked.FirstOrDefault();
            if (chosen == null)
            {
                chosen = subfolders.FirstOrDefault(f => f.Mark == FolderMark.Unknown
                    && (OutlookFolderNames.Matches(f.Name, groupName)
                        || OutlookFolderNames.Matches(f.Name, groupName + " (Nextcloud)")));
                if (chosen != null)
                {
                    DiagnosticsLogger.Log(
                        LogCategories.Core,
                        "CardDAV group folder marker unreadable; reusing same-named folder instead of creating a duplicate (group="
                        + groupName + ").");
                }
            }

            Outlook.MAPIFolder result = null;
            if (chosen != null)
            {
                result = TryGetFolder(session, chosen.EntryId, storeId);
            }
            if (result == null)
            {
                result = CreateManagedGroupFolder(parent, groupName, subfolders);
                owned = result != null;
                if (result != null)
                {
                    subfolders.Add(new SubfolderInfo
                    {
                        EntryId = result.EntryID,
                        Name = result.Name,
                        Mark = FolderMark.Managed,
                        GroupName = groupName
                    });
                }
            }
            if (result == null) return null;

            if (owned)
            {
                state.Set(groupName, result.EntryID);
                RemoveDuplicateGroupFolders(session, storeId, groupName, result.EntryID, marked, subfolders);
            }
            return result;
        }

        private static Outlook.MAPIFolder CreateManagedGroupFolder(
            Outlook.MAPIFolder parent,
            string groupName,
            IEnumerable<SubfolderInfo> subfolders)
        {
            string folderName = SanitizeFolderName(groupName);
            if (subfolders.Any(f => OutlookFolderNames.Matches(f.Name, folderName)))
            {
                folderName = SanitizeFolderName(groupName + " (Nextcloud)");
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "CardDAV group folder name collision; using alternate folder (group=" + groupName
                    + ", folder=" + folderName + ").");
            }
            if (subfolders.Any(f => OutlookFolderNames.Matches(f.Name, folderName)))
            {
                // Both names are taken by folders we cannot identify; skip rather than fail or duplicate.
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "CardDAV group folder skipped because its names are taken (group=" + groupName + ").");
                return null;
            }

            Outlook.Folders folders = null;
            try
            {
                folders = parent.Folders;
                Outlook.MAPIFolder created = folders.Add(folderName, Outlook.OlDefaultFolders.olFolderContacts);
                WriteFolderProperty(created, ManagedFolderProperty, "1");
                WriteFolderProperty(created, ManagedFolderNameProperty, groupName);
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "CardDAV group folder created (group=" + groupName + ", folder=" + folderName + ").");
                return created;
            }
            finally
            {
                ComInteropScope.TryRelease(folders, LogCategories.Core, "Failed to release folders collection after CardDAV group folder creation.");
            }
        }

        /// <summary>
        /// Deletes extra folders of the same group, but only when they hold nothing except managed copies.
        /// Deleted folders go to "Deleted Items".
        /// </summary>
        private static void RemoveDuplicateGroupFolders(
            Outlook.NameSpace session,
            string storeId,
            string groupName,
            string keepEntryId,
            IEnumerable<SubfolderInfo> marked,
            List<SubfolderInfo> subfolders)
        {
            foreach (SubfolderInfo duplicate in marked.Where(f => !string.Equals(f.EntryId, keepEntryId, StringComparison.Ordinal)).ToList())
            {
                Outlook.MAPIFolder folder = TryGetFolder(session, duplicate.EntryId, storeId);
                if (folder == null) continue;
                try
                {
                    if (!ContainsOnlyManagedItems(folder))
                    {
                        DiagnosticsLogger.Log(
                            LogCategories.Core,
                            "CardDAV duplicate group folder kept because it contains own items (group=" + groupName + ").");
                        continue;
                    }
                    folder.Delete();
                    subfolders.Remove(duplicate);
                    DiagnosticsLogger.Log(
                        LogCategories.Core,
                        "CardDAV duplicate group folder removed (group=" + groupName + ", folder=" + duplicate.Name + ").");
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to remove duplicate CardDAV group folder.", ex);
                }
                finally
                {
                    ComInteropScope.TryRelease(folder, LogCategories.Core, "Failed to release duplicate CardDAV group folder.");
                }
            }
        }

        private static bool ContainsOnlyManagedItems(Outlook.MAPIFolder folder)
        {
            Outlook.Folders children = null;
            Outlook.Items items = null;
            try
            {
                children = folder.Folders;
                if (children != null && children.Count > 0) return false;
                items = folder.Items;
                int count = items != null ? items.Count : 0;
                for (int index = 1; index <= count; index++)
                {
                    object raw = null;
                    try
                    {
                        raw = items[index];
                        var contact = raw as Outlook.ContactItem;
                        var list = raw as Outlook.DistListItem;
                        bool managed = contact != null
                            ? string.Equals(ReadContactProperty(contact, GroupCopyPropertyName), "1", StringComparison.Ordinal)
                            : list != null
                                && string.Equals(ReadDistributionListProperty(list, ManagedDistributionListPropertyName), "1", StringComparison.Ordinal);
                        if (!managed) return false;
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(raw, LogCategories.Core, "Failed to release item while checking duplicate CardDAV group folder.");
                    }
                }
                return true;
            }
            finally
            {
                ComInteropScope.TryRelease(items, LogCategories.Core, "Failed to release items while checking duplicate CardDAV group folder.");
                ComInteropScope.TryRelease(children, LogCategories.Core, "Failed to release subfolders while checking duplicate CardDAV group folder.");
            }
        }

        private static Outlook.MAPIFolder TryGetFolder(Outlook.NameSpace session, string entryId, string storeId)
        {
            if (string.IsNullOrWhiteSpace(entryId)) return null;
            try
            {
                return session.GetFolderFromID(entryId, storeId);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return null;
            }
        }

        private static void DeleteObsoleteManagedFolders(
            Outlook.NameSpace session,
            string storeId,
            List<SubfolderInfo> subfolders,
            ISet<string> activeGroups,
            GroupFolderState state)
        {
            foreach (SubfolderInfo info in subfolders.ToList())
            {
                string groupName = state.GroupFor(info.EntryId);
                if (groupName.Length == 0 && info.Mark == FolderMark.Managed) groupName = info.GroupName;
                // Without a known group name the folder cannot be proven obsolete, so it stays.
                if (groupName.Length == 0 || activeGroups.Contains(groupName)) continue;

                Outlook.MAPIFolder folder = TryGetFolder(session, info.EntryId, storeId);
                if (folder == null) continue;
                try
                {
                    folder.Delete();
                    subfolders.Remove(info);
                    state.Remove(groupName);
                    DiagnosticsLogger.Log(
                        LogCategories.Core,
                        "CardDAV obsolete group folder removed (group=" + groupName + ").");
                }
                finally
                {
                    ComInteropScope.TryRelease(folder, LogCategories.Core, "Failed to release obsolete CardDAV group folder.");
                }
            }
        }

        /// <summary>
        /// Removes the add-in's group folders for a full rebuild. A folder that also holds the user's own
        /// items keeps them and only loses the managed copies. Unreadable markers are left alone.
        /// </summary>
        internal static int RemoveManagedFolders(Outlook.NameSpace session, Outlook.MAPIFolder parent)
        {
            GroupFolderState state = GroupFolderState.Load(parent.EntryID);
            string storeId = parent.StoreID;
            int removed = 0;
            foreach (SubfolderInfo info in ScanSubfolders(parent))
            {
                if (info.Mark != FolderMark.Managed && state.GroupFor(info.EntryId).Length == 0) continue;
                Outlook.MAPIFolder folder = TryGetFolder(session, info.EntryId, storeId);
                if (folder == null) continue;
                try
                {
                    if (ContainsOnlyManagedItems(folder))
                    {
                        folder.Delete();
                        removed++;
                    }
                    else
                    {
                        ClearManagedCopies(folder);
                        DeleteManagedDistributionLists(folder);
                        WriteFolderProperty(folder, ManagedFolderSignatureProperty, string.Empty);
                        DiagnosticsLogger.Log(
                            LogCategories.Core,
                            "CardDAV group folder kept during reset because it contains own items (folder=" + info.Name + ").");
                    }
                }
                finally
                {
                    ComInteropScope.TryRelease(folder, LogCategories.Core, "Failed to release CardDAV group folder during reset.");
                }
            }
            state.Clear();
            state.Save();
            return removed;
        }

        internal static bool IsManagedGroupFolder(Outlook.MAPIFolder folder)
        {
            string groupName;
            return ReadFolderMark(folder, out groupName) == FolderMark.Managed;
        }

        private static FolderMark ReadFolderMark(Outlook.MAPIFolder folder, out string groupName)
        {
            groupName = string.Empty;
            string managed;
            if (!TryReadFolderProperty(folder, ManagedFolderProperty, out managed)) return FolderMark.Unknown;
            if (!string.Equals(managed, "1", StringComparison.Ordinal)) return FolderMark.Unmanaged;
            string name;
            if (!TryReadFolderProperty(folder, ManagedFolderNameProperty, out name) || string.IsNullOrWhiteSpace(name))
            {
                return FolderMark.Unknown;
            }
            groupName = name.Trim();
            return FolderMark.Managed;
        }

        private static string ReadFolderProperty(Outlook.MAPIFolder folder, string schemaName)
        {
            string value;
            return TryReadFolderProperty(folder, schemaName, out value) ? value : string.Empty;
        }

        /// <summary>
        /// Returns false when Outlook failed to read the property; a missing property is a successful empty read.
        /// </summary>
        private static bool TryReadFolderProperty(Outlook.MAPIFolder folder, string schemaName, out string value)
        {
            const int MapiNotFound = unchecked((int)0x8004010F);
            value = string.Empty;
            Outlook.PropertyAccessor accessor = null;
            try
            {
                accessor = folder != null ? folder.PropertyAccessor : null;
                if (accessor == null) return false;
                object raw = accessor.GetProperty(schemaName);
                value = raw != null ? Convert.ToString(raw) ?? string.Empty : string.Empty;
                return true;
            }
            catch (Exception ex)
            {
                var com = ex as System.Runtime.InteropServices.COMException;
                if (com != null && com.ErrorCode == MapiNotFound)
                {
                    return true;
                }
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "CardDAV group folder property could not be read (property=" + schemaName.Substring(schemaName.LastIndexOf('/') + 1) + ").",
                    ex);
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(accessor, LogCategories.Core, "Failed to release CardDAV group folder property accessor.");
            }
        }

        /// <summary>
        /// Remembers which folder belongs to which group, so identification does not depend on folder properties.
        /// </summary>
        private sealed class GroupFolderState
        {
            private const string FileName = "carddav-group-folders.xml";
            private readonly string _parentEntryId;
            private readonly Dictionary<string, string> _entryIds = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            private bool _dirty;

            private GroupFolderState(string parentEntryId)
            {
                _parentEntryId = parentEntryId ?? string.Empty;
            }

            internal static GroupFolderState Load(string parentEntryId)
            {
                var state = new GroupFolderState(parentEntryId);
                try
                {
                    string path = GetPath();
                    if (!File.Exists(path)) return state;
                    XElement root = XDocument.Load(path).Root;
                    // Another default contacts folder (new profile or store) starts from the folder markers.
                    if (root == null || !string.Equals((string)root.Attribute("parent"), state._parentEntryId, StringComparison.Ordinal))
                    {
                        return state;
                    }
                    foreach (XElement folder in root.Elements("Folder"))
                    {
                        string group = ((string)folder.Attribute("group") ?? string.Empty).Trim();
                        string entryId = ((string)folder.Attribute("entryId") ?? string.Empty).Trim();
                        if (group.Length > 0 && entryId.Length > 0) state._entryIds[group] = entryId;
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to load CardDAV group folder state.", ex);
                }
                return state;
            }

            internal string Get(string groupName)
            {
                string entryId;
                return _entryIds.TryGetValue(groupName ?? string.Empty, out entryId) ? entryId : string.Empty;
            }

            internal string GroupFor(string entryId)
            {
                return _entryIds.Where(pair => string.Equals(pair.Value, entryId, StringComparison.Ordinal))
                    .Select(pair => pair.Key)
                    .FirstOrDefault() ?? string.Empty;
            }

            internal void Set(string groupName, string entryId)
            {
                if (string.IsNullOrWhiteSpace(groupName) || string.IsNullOrWhiteSpace(entryId)) return;
                if (string.Equals(Get(groupName), entryId, StringComparison.Ordinal)) return;
                _entryIds[groupName] = entryId;
                _dirty = true;
            }

            internal void Clear()
            {
                _entryIds.Clear();
                _dirty = true;
            }

            internal void Remove(string groupName)
            {
                if (_entryIds.Remove(groupName ?? string.Empty)) _dirty = true;
            }

            internal void Save()
            {
                if (!_dirty) return;
                try
                {
                    var root = new XElement("CardDavGroupFolders", new XAttribute("parent", _parentEntryId));
                    foreach (KeyValuePair<string, string> pair in _entryIds.OrderBy(p => p.Key, StringComparer.OrdinalIgnoreCase))
                    {
                        root.Add(new XElement("Folder", new XAttribute("group", pair.Key), new XAttribute("entryId", pair.Value)));
                    }
                    new XDocument(root).Save(GetPath());
                    _dirty = false;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to save CardDAV group folder state.", ex);
                }
            }

            private static string GetPath()
            {
                return Path.Combine(AppDataPaths.EnsureLocalRootDirectory(), FileName);
            }
        }

        private static void WriteFolderProperty(Outlook.MAPIFolder folder, string schemaName, string value)
        {
            Outlook.PropertyAccessor accessor = null;
            try
            {
                accessor = folder.PropertyAccessor;
                accessor.SetProperty(schemaName, value ?? string.Empty);
            }
            finally
            {
                ComInteropScope.TryRelease(accessor, LogCategories.Core, "Failed to release CardDAV group folder property accessor after write.");
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
                ComInteropScope.TryRelease(property, LogCategories.Core, "Failed to release CardDAV group folder contact property.");
                ComInteropScope.TryRelease(properties, LogCategories.Core, "Failed to release CardDAV group folder contact properties.");
            }
        }

        private static string ReadDistributionListProperty(Outlook.DistListItem item, string name)
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
                ComInteropScope.TryRelease(property, LogCategories.Core, "Failed to release CardDAV distribution-list property.");
                ComInteropScope.TryRelease(properties, LogCategories.Core, "Failed to release CardDAV distribution-list properties.");
            }
        }

        private static void WriteContactProperty(Outlook.ContactItem item, string name, string value)
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
                ComInteropScope.TryRelease(property, LogCategories.Core, "Failed to release CardDAV group folder contact property after write.");
                ComInteropScope.TryRelease(properties, LogCategories.Core, "Failed to release CardDAV group folder contact properties after write.");
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
                if (!preferences.IncludesAddressBook(addressBook))
                {
                    continue;
                }

                XDocument document = SendCategoriesQuery(configuration, addressBook);
                foreach (XElement responseElement in document.Descendants(Dav + "response"))
                {
                    XElement prop = responseElement
                        .Elements(Dav + "propstat")
                        .Where(p => IsSuccessfulStatus((string)p.Element(Dav + "status")))
                        .Select(p => p.Element(Dav + "prop"))
                        .FirstOrDefault(p => p != null);
                    if (prop == null) continue;

                    string requestHref = ((string)responseElement.Element(Dav + "href") ?? string.Empty).Trim();
                    string href = ResolveUri(addressBook.Href, requestHref);
                    string vcard = (string)prop.Element(CardDav + "address-data");
                    if (string.IsNullOrWhiteSpace(href) || string.IsNullOrWhiteSpace(vcard)) continue;
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
                throw new InvalidOperationException("CardDAV address book is missing for group folder sync.");
            }

            string body = "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                + "<card:addressbook-query xmlns:d=\"DAV:\" xmlns:card=\"urn:ietf:params:xml:ns:carddav\">"
                + "<d:prop><card:address-data content-type=\"text/vcard\">"
                + "<card:prop name=\"CATEGORIES\"/>"
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
                throw new InvalidOperationException("CardDAV group folder query failed for " + addressBook.Href + ".");
            }

            XDocument document = XDocument.Parse(response.ResponseText ?? string.Empty);
            if (document.Root == null || document.Root.Name != Dav + "multistatus"
                || document.Descendants(Dav + "error").Any()
                || document.Descendants(Dav + "status").Any(status => !IsSuccessfulStatus(status.Value)))
            {
                throw new InvalidOperationException("Incomplete CardDAV group folder response for " + addressBook.Href + ".");
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
                if (colon <= 0) continue;
                string field = line.Substring(0, colon);
                string name = field.Split(';')[0];
                if (!string.Equals(name, "CATEGORIES", StringComparison.OrdinalIgnoreCase)) continue;

                foreach (string category in SplitCategoryValues(line.Substring(colon + 1)))
                {
                    string value = (category ?? string.Empty).Trim();
                    if (!string.IsNullOrWhiteSpace(value)
                        && !result.Any(existing => string.Equals(existing, value, StringComparison.OrdinalIgnoreCase)))
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
                    if (c == 'n' || c == 'N') current.Append(Environment.NewLine);
                    else current.Append(c);
                    escaped = false;
                }
                else if (c == '\\') escaped = true;
                else if (c == ',')
                {
                    result.Add(current.ToString());
                    current.Clear();
                }
                else current.Append(c);
            }
            if (escaped) current.Append('\\');
            result.Add(current.ToString());
            return result;
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

        private static string SanitizeFolderName(string value)
        {
            string result = (value ?? string.Empty).Trim();
            foreach (char c in new[] { '\\', '/', ':', '*', '?', '"', '<', '>', '|' })
            {
                result = result.Replace(c, '-');
            }
            return string.IsNullOrWhiteSpace(result) ? "Nextcloud-Gruppe" : result;
        }
    }
}
