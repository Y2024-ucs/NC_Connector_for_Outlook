// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class CardDavContactsResetResult
    {
        internal int Contacts { get; set; }
        internal int DistributionLists { get; set; }
        internal int GroupFolders { get; set; }
    }

    /// <summary>
    /// Removes everything the read-only contact sync created in Outlook, so the next sync rebuilds it.
    /// Contacts without add-in markers and the two-way customer folder are never touched.
    /// Deleted items go to "Deleted Items".
    /// </summary>
    internal static class CardDavContactsReset
    {
        private const int SliceBudgetMilliseconds = 40;

        internal static async Task<CardDavContactsResetResult> RunAsync(
            Outlook.Application application,
            Func<Action, Task> runOnUi,
            Func<Task> yieldUi)
        {
            if (application == null) throw new ArgumentNullException("application");
            if (runOnUi == null) throw new ArgumentNullException("runOnUi");
            if (yieldUi == null) yieldUi = () => Task.Delay(20);

            var result = new CardDavContactsResetResult();
            List<string> contactIds = null;
            string storeId = null;
            await runOnUi(() =>
            {
                Outlook.NameSpace session = null;
                Outlook.MAPIFolder root = null;
                try
                {
                    session = application.Session;
                    root = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                    storeId = root.StoreID;
                    contactIds = CardDavReadOnlySync.CollectImportedContactIds(root);
                }
                finally
                {
                    ComInteropScope.TryRelease(root, LogCategories.Core, "Failed to release contacts folder during CardDAV reset scan.");
                    ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session during CardDAV reset scan.");
                }
            }).ConfigureAwait(false);

            int index = 0;
            while (index < contactIds.Count)
            {
                await runOnUi(() =>
                {
                    Outlook.NameSpace session = null;
                    try
                    {
                        session = application.Session;
                        var stopwatch = Stopwatch.StartNew();
                        do
                        {
                            Outlook.ContactItem contact = CardDavReadOnlySync.TryGetContactById(session, contactIds[index], storeId);
                            index++;
                            if (contact == null) continue;
                            try
                            {
                                contact.Delete();
                                result.Contacts++;
                            }
                            finally
                            {
                                ComInteropScope.TryRelease(contact, LogCategories.Core, "Failed to release contact during CardDAV reset.");
                            }
                        }
                        while (index < contactIds.Count && stopwatch.ElapsedMilliseconds < SliceBudgetMilliseconds);
                    }
                    finally
                    {
                        ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session during CardDAV reset.");
                    }
                }).ConfigureAwait(false);
                if (index < contactIds.Count) await yieldUi().ConfigureAwait(false);
            }

            await runOnUi(() =>
            {
                Outlook.NameSpace session = null;
                Outlook.MAPIFolder root = null;
                try
                {
                    session = application.Session;
                    root = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                    result.DistributionLists = CardDavContactGroupSync.RemoveManagedGroups(root);
                    result.GroupFolders = CardDavContactFolderSync.RemoveManagedFolders(session, root);
                }
                finally
                {
                    ComInteropScope.TryRelease(root, LogCategories.Core, "Failed to release contacts folder during CardDAV group reset.");
                    ComInteropScope.TryRelease(session, LogCategories.Core, "Failed to release Outlook session during CardDAV group reset.");
                }
            }).ConfigureAwait(false);

            CardDavReadOnlySync.ClearKnownEtags();
            DiagnosticsLogger.Log(
                LogCategories.Core,
                "CardDAV contact sync reset (contacts=" + result.Contacts
                + ", distributionLists=" + result.DistributionLists
                + ", groupFolders=" + result.GroupFolders + ").");
            return result;
        }
    }
}
