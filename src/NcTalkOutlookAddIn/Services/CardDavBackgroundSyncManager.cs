// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Threading;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Services
{
    internal static class CardDavBackgroundSyncManager
    {
        private static readonly object SyncRoot = new object();
        private static readonly TimeSpan Interval = TimeSpan.FromMinutes(15);
        private static Timer _timer;
        private static TalkServiceConfiguration _configuration;
        private static Outlook.Application _outlookApplication;
        private static OutlookUiSynchronizationContext _uiContext;
        private static int _running;

        internal static void EnsureStarted(
            TalkServiceConfiguration configuration,
            Outlook.Application outlookApplication)
        {
            if (configuration == null || outlookApplication == null)
            {
                return;
            }

            lock (SyncRoot)
            {
                _configuration = configuration;
                _outlookApplication = outlookApplication;

                if (_uiContext == null)
                {
                    _uiContext = new OutlookUiSynchronizationContext();
                }

                if (_timer != null)
                {
                    return;
                }

                _timer = new Timer(OnTimer, null, Interval, Interval);
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "CardDAV background sync enabled (intervalMinutes=15).");
            }
        }

        private static void OnTimer(object state)
        {
            if (Interlocked.CompareExchange(ref _running, 1, 0) != 0)
            {
                return;
            }

            try
            {
                CardDavSyncPreferences preferences = CardDavSyncPreferences.Load();
                if (!preferences.Enabled
                    || (!preferences.SyncCompanyDirectory && !preferences.SyncPersonalContacts))
                {
                    Interlocked.Exchange(ref _running, 0);
                    return;
                }

                TalkServiceConfiguration configuration;
                Outlook.Application outlookApplication;
                OutlookUiSynchronizationContext uiContext;
                lock (SyncRoot)
                {
                    configuration = _configuration;
                    outlookApplication = _outlookApplication;
                    uiContext = _uiContext;
                }

                if (configuration == null
                    || !configuration.IsComplete()
                    || outlookApplication == null
                    || uiContext == null)
                {
                    Interlocked.Exchange(ref _running, 0);
                    return;
                }

                var sync = new CardDavReadOnlySync(configuration);
                IList<CardDavAddressBook> addressBooks =
                    new DavDiscoveryService(configuration).DiscoverAddressBooks();
                var contacts = new List<CardDavContactRecord>();

                foreach (CardDavAddressBook addressBook in addressBooks)
                {
                    bool isSystem = IsSystemAddressBook(addressBook);
                    if ((isSystem && !preferences.SyncCompanyDirectory)
                        || (!isSystem && !preferences.SyncPersonalContacts))
                    {
                        continue;
                    }

                    contacts.AddRange(sync.DownloadContacts(addressBook));
                }

                uiContext.Post(
                    delegate
                    {
                        try
                        {
                            int count = sync.ImportIntoOutlook(outlookApplication, contacts);
                            DiagnosticsLogger.Log(
                                LogCategories.Core,
                                "CardDAV background sync completed (contacts=" + count + ").");
                        }
                        catch (Exception ex)
                        {
                            DiagnosticsLogger.LogException(
                                LogCategories.Core,
                                "CardDAV background Outlook import failed.",
                                ex);
                        }
                        finally
                        {
                            Interlocked.Exchange(ref _running, 0);
                        }
                    },
                    null);
            }
            catch (Exception ex)
            {
                Interlocked.Exchange(ref _running, 0);
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "CardDAV background sync failed.",
                    ex);
            }
        }

        private static bool IsSystemAddressBook(CardDavAddressBook addressBook)
        {
            return addressBook != null
                && !string.IsNullOrWhiteSpace(addressBook.Href)
                && addressBook.Href.IndexOf(
                    "z-server-generated--system",
                    StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
