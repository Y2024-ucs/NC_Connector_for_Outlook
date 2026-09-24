// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
    // Binds only selected calendar appointments.
    // This keeps calendar-view deletion working after restart without opening every
    // Outlook store, folder, or item at startup.
    public sealed partial class NextcloudTalkAddIn
    {
        private void OnExplorerSelectionChanged(string explorerKey)
        {
            Outlook.Explorer explorer;
            if (string.IsNullOrWhiteSpace(explorerKey)
                || !_hookedExplorers.TryGetValue(explorerKey, out explorer)
                || explorer == null)
            {
                return;
            }

            Outlook.MAPIFolder folder = null;
            Outlook.Selection selection = null;
            try
            {
                folder = explorer.CurrentFolder;
                if (folder == null
                    || folder.DefaultItemType != Outlook.OlItemType.olAppointmentItem)
                {
                    return;
                }

                selection = explorer.Selection;
                int count = selection != null ? selection.Count : 0;
                for (int index = 1; index <= count; index++)
                {
                    object selectedItem = null;
                    bool selectionReferenceRetained = false;
                    try
                    {
                        selectedItem = selection[index];
                        Outlook.AppointmentItem appointment =
                            selectedItem as Outlook.AppointmentItem;
                        if (appointment != null)
                        {
                            EnsureSubscriptionForSelectedAppointment(
                                appointment,
                                out selectionReferenceRetained);
                        }
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            LogCategories.Talk,
                            "Failed to bind a selected calendar appointment.",
                            ex);
                    }
                    finally
                    {
                        if (!selectionReferenceRetained)
                        {
                            ComInteropScope.TryRelease(
                                selectedItem,
                                LogCategories.Talk,
                                "Failed to release an unbound Explorer selection item.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Talk,
                    "Failed to process the Explorer calendar selection.",
                    ex);
            }
            finally
            {
                ComInteropScope.TryRelease(
                    selection,
                    LogCategories.Talk,
                    "Failed to release the Explorer Selection COM object.");
                ComInteropScope.TryRelease(
                    folder,
                    LogCategories.Talk,
                    "Failed to release Explorer.CurrentFolder after calendar selection processing.");
            }
        }
    }
}
