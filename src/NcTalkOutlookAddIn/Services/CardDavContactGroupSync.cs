// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Reflection;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Services
{
    internal static class CardDavContactGroupSync
    {
        internal const string CategoriesPropertyName = "NC-CardDAV-CATEGORIES";
        private const string HrefPropertyName = "NC-CardDAV-HREF";
        private const string ManagedGroupPropertyName = "NC-CardDAV-GROUP";
        private const string ManagedGroupNamePropertyName = "NC-CardDAV-GROUP-NAME";

        private sealed class GroupMember
        {
            internal string DisplayName { get; set; }
            internal string EmailAddress { get; set; }
        }

        internal static int Reconcile(Outlook.Application outlookApplication)
        {
            if (outlookApplication == null)
            {
                return 0;
            }

            Outlook.NameSpace session = null;
            Outlook.MAPIFolder defaultContacts = null;
            Outlook.Folders folders = null;
            int groupCount = 0;
            try
            {
                session = outlookApplication.Session;
                defaultContacts = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                groupCount += ReconcileFolder(session, defaultContacts);

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
                            groupCount += ReconcileFolder(session, folder);
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

        private static int ReconcileFolder(Outlook.NameSpace session, Outlook.MAPIFolder folder)
        {
            if (session == null || folder == null)
            {
                return 0;
            }

            Dictionary<string, List<GroupMember>> groups = LoadManagedContactGroups(folder);
            DeleteManagedGroups(folder);

            int created = 0;
            foreach (KeyValuePair<string, List<GroupMember>> pair in groups)
            {
                if (string.IsNullOrWhiteSpace(pair.Key) || pair.Value == null || pair.Value.Count == 0)
                {
                    continue;
                }

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

                    list.DLName = pair.Key;
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
                                    "CardDAV contact group member could not be resolved (group="
                                    + pair.Key + ", member=" + (member.DisplayName ?? string.Empty) + ").");
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

        private static Dictionary<string, List<GroupMember>> LoadManagedContactGroups(Outlook.MAPIFolder folder)
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
                        if (string.IsNullOrWhiteSpace(href))
                        {
                            continue;
                        }

                        string serializedCategories = ReadContactProperty(contact, CategoriesPropertyName);
                        string[] categories = (serializedCategories ?? string.Empty)
                            .Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
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
                        if (list == null)
                        {
                            continue;
                        }
                        if (string.Equals(ReadGroupProperty(list, ManagedGroupPropertyName), "1", StringComparison.Ordinal))
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

        internal static string SerializeCategories(IEnumerable<string> categories)
        {
            if (categories == null)
            {
                return string.Empty;
            }
            var values = new List<string>();
            foreach (string raw in categories)
            {
                string value = (raw ?? string.Empty).Trim();
                if (!string.IsNullOrWhiteSpace(value) && !values.Contains(value))
                {
                    values.Add(value);
                }
            }
            return string.Join("\n", values.ToArray());
        }

        internal static bool HasCategoriesProperty(Outlook.ContactItem item)
        {
            Outlook.UserProperties properties = null;
            Outlook.UserProperty property = null;
            try
            {
                properties = item != null ? item.UserProperties : null;
                property = properties != null ? properties.Find(CategoriesPropertyName, true) : null;
                return property != null;
            }
            finally
            {
                ComInteropScope.TryRelease(property, LogCategories.Core, "Failed to release CardDAV categories marker property.");
                ComInteropScope.TryRelease(properties, LogCategories.Core, "Failed to release CardDAV categories marker properties.");
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
    }
}
