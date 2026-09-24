// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
        // Outlook hook lifecycle for inspector/application events and compose-tracking entry points.
    public sealed partial class NextcloudTalkAddIn
    {
        private void EnsureInspectorHook()
        {
            if (_outlookApplication == null || _inspectors != null)
            {
                return;
            }
            try
            {
                _inspectors = _outlookApplication.Inspectors;
                if (_inspectors != null)
                {
                    _inspectors.NewInspector += OnNewInspector;
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to hook Inspectors.NewInspector.", ex);
                _inspectors = null;
            }
        }

        private void EnsureApplicationHook()
        {
            if (_outlookApplication == null)
            {
                return;
            }
            try
            {
                EnsureExplorerHooks();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to hook Explorer lifecycle.", ex);
            }
        }

        private void UnhookApplication()
        {
            UnhookExplorerHooks();
        }

        private void EnsureExplorerHooks()
        {
            if (_outlookApplication == null)
            {
                return;
            }
            try
            {
                if (_explorers == null)
                {
                    _explorers = _outlookApplication.Explorers;
                    _explorersEvents = _explorers as Outlook.ExplorersEvents_Event;
                    if (_explorersEvents != null)
                    {
                        _explorersEvents.NewExplorer += OnNewExplorer;
                    }

                    int count = _explorers != null ? _explorers.Count : 0;
                    for (int i = 1; i <= count; i++)
                    {
                        Outlook.Explorer explorer = null;
                        try
                        {
                            explorer = _explorers[i];
                            if (!HookExplorer(explorer))
                            {
                                ComInteropScope.TryRelease(
                                    explorer,
                                    LogCategories.Core,
                                    "Failed to release an untracked existing Explorer.");
                            }
                        }
                        catch (Exception ex)
                        {
                            DiagnosticsLogger.LogException(LogCategories.Core, "Failed to hook existing Explorer lifecycle.", ex);
                            ComInteropScope.TryRelease(explorer, LogCategories.Core, "Failed to release Explorer after hook failure.");
                        }
                    }
                }

                Outlook.Explorer activeExplorer = null;
                try
                {
                    activeExplorer = _outlookApplication.ActiveExplorer();
                    if (!HookExplorer(activeExplorer))
                    {
                        ComInteropScope.TryRelease(
                            activeExplorer,
                            LogCategories.Core,
                            "Failed to release an untracked ActiveExplorer.");
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to hook active Explorer lifecycle.", ex);
                    ComInteropScope.TryRelease(activeExplorer, LogCategories.Core, "Failed to release ActiveExplorer after hook failure.");
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to initialize Explorer hooks.", ex);
            }
        }

        private bool HookExplorer(Outlook.Explorer explorer)
        {
            if (explorer == null)
            {
                return false;
            }
            string explorerKey = ComInteropScope.ResolveIdentityKey(explorer, LogCategories.Core, "Explorer");
            if (string.IsNullOrWhiteSpace(explorerKey) || _hookedExplorerEvents.ContainsKey(explorerKey))
            {
                return false;
            }

            var explorerEvents = explorer as Outlook.ExplorerEvents_10_Event;
            if (explorerEvents == null)
            {
                return false;
            }

            Outlook.ExplorerEvents_10_InlineResponseEventHandler inlineResponseHandler =
                item => OnExplorerInlineResponse(explorerKey, item);
            Outlook.ExplorerEvents_10_InlineResponseCloseEventHandler inlineResponseCloseHandler =
                () => OnExplorerInlineResponseClose(explorerKey);
            Outlook.ExplorerEvents_10_SelectionChangeEventHandler selectionChangeHandler =
                () => OnExplorerSelectionChanged(explorerKey);
            Outlook.ExplorerEvents_10_FolderSwitchEventHandler folderSwitchHandler =
                () => OnExplorerFolderSwitched(explorerKey);

            explorerEvents.InlineResponse += inlineResponseHandler;
            try
            {
                explorerEvents.InlineResponseClose += inlineResponseCloseHandler;
            }
            catch (Exception ex)
            {
                try
                {
                    explorerEvents.InlineResponse -= inlineResponseHandler;
                }
                catch (Exception rollbackEx)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Core,
                        "Failed to roll back Explorer.InlineResponse hook (explorerKey=" + explorerKey + ").",
                        rollbackEx);
                }
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "Failed to hook Explorer.InlineResponseClose (explorerKey=" + explorerKey + ").",
                    ex);
                return false;
            }
            try
            {
                explorerEvents.FolderSwitch += folderSwitchHandler;
            }
            catch (Exception ex)
            {
                try
                {
                    explorerEvents.InlineResponseClose -= inlineResponseCloseHandler;
                    explorerEvents.InlineResponse -= inlineResponseHandler;
                }
                catch (Exception rollbackEx)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Core,
                        "Failed to roll back Explorer hooks after FolderSwitch failure (explorerKey=" + explorerKey + ").",
                        rollbackEx);
                }
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "Failed to hook Explorer.FolderSwitch (explorerKey=" + explorerKey + ").",
                    ex);
                return false;
            }

            _hookedExplorerEvents[explorerKey] = explorerEvents;
            _inlineResponseHandlers[explorerKey] = inlineResponseHandler;
            _inlineResponseCloseHandlers[explorerKey] = inlineResponseCloseHandler;
            _explorerFolderSwitchHandlers[explorerKey] = folderSwitchHandler;
            _hookedExplorers[explorerKey] = explorer;

            UpdateExplorerSelectionHook(explorerKey, selectionChangeHandler);

            LogCore(
                "Explorer lifecycle hooked (explorerKey="
                + explorerKey
                + ", calendarSelection="
                + _explorerSelectionChangeHandlers.ContainsKey(explorerKey)
                + ").");
            return true;
        }

        private void OnExplorerFolderSwitched(string explorerKey)
        {
            try
            {
                UpdateExplorerSelectionHook(explorerKey, null);
                RefreshCardDavRibbon();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "Failed to process Explorer.FolderSwitch (explorerKey=" + (explorerKey ?? string.Empty) + ").",
                    ex);
            }
        }

        private void UpdateExplorerSelectionHook(
            string explorerKey,
            Outlook.ExplorerEvents_10_SelectionChangeEventHandler preferredHandler)
        {
            Outlook.Explorer explorer;
            Outlook.ExplorerEvents_10_Event explorerEvents;
            if (string.IsNullOrWhiteSpace(explorerKey)
                || !_hookedExplorers.TryGetValue(explorerKey, out explorer)
                || explorer == null
                || !_hookedExplorerEvents.TryGetValue(explorerKey, out explorerEvents)
                || explorerEvents == null)
            {
                return;
            }

            bool shouldHook = IsCalendarExplorer(explorer);
            Outlook.ExplorerEvents_10_SelectionChangeEventHandler existingHandler;
            bool isHooked = _explorerSelectionChangeHandlers.TryGetValue(
                explorerKey,
                out existingHandler);

            if (shouldHook && !isHooked)
            {
                Outlook.ExplorerEvents_10_SelectionChangeEventHandler handler =
                    preferredHandler ?? (() => OnExplorerSelectionChanged(explorerKey));
                try
                {
                    explorerEvents.SelectionChange += handler;
                    _explorerSelectionChangeHandlers[explorerKey] = handler;
                    OnExplorerSelectionChanged(explorerKey);
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Core,
                        "Failed to hook calendar Explorer.SelectionChange (explorerKey=" + explorerKey + ").",
                        ex);
                }
                return;
            }

            if (!shouldHook && isHooked)
            {
                try
                {
                    explorerEvents.SelectionChange -= existingHandler;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Core,
                        "Failed to unhook non-calendar Explorer.SelectionChange (explorerKey=" + explorerKey + ").",
                        ex);
                }
                _explorerSelectionChangeHandlers.Remove(explorerKey);
            }
        }

        private static bool IsCalendarExplorer(Outlook.Explorer explorer)
        {
            if (explorer == null)
            {
                return false;
            }

            Outlook.MAPIFolder folder = null;
            try
            {
                folder = explorer.CurrentFolder;
                return folder != null
                    && folder.DefaultItemType == Outlook.OlItemType.olAppointmentItem;
            }
            catch
            {
                return false;
            }
            finally
            {
                ComInteropScope.TryRelease(
                    folder,
                    LogCategories.Core,
                    "Failed to release Explorer.CurrentFolder while checking calendar selection hook.");
            }
        }

        private void UnhookExplorerHooks()
        {
            foreach (var pair in _hookedExplorerEvents)
            {
                Outlook.ExplorerEvents_10_FolderSwitchEventHandler folderSwitchHandler;
                if (_explorerFolderSwitchHandlers.TryGetValue(pair.Key, out folderSwitchHandler))
                {
                    try
                    {
                        pair.Value.FolderSwitch -= folderSwitchHandler;
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            LogCategories.Core,
                            "Failed to unhook Explorer.FolderSwitch (explorerKey=" + pair.Key + ").",
                            ex);
                    }
                }
                Outlook.ExplorerEvents_10_InlineResponseEventHandler inlineResponseHandler;
                if (_inlineResponseHandlers.TryGetValue(pair.Key, out inlineResponseHandler))
                {
                    try
                    {
                        pair.Value.InlineResponse -= inlineResponseHandler;
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            LogCategories.Core,
                            "Failed to unhook Explorer.InlineResponse (explorerKey=" + pair.Key + ").",
                            ex);
                    }
                }

                Outlook.ExplorerEvents_10_InlineResponseCloseEventHandler inlineResponseCloseHandler;
                if (_inlineResponseCloseHandlers.TryGetValue(pair.Key, out inlineResponseCloseHandler))
                {
                    try
                    {
                        pair.Value.InlineResponseClose -= inlineResponseCloseHandler;
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            LogCategories.Core,
                            "Failed to unhook Explorer.InlineResponseClose (explorerKey=" + pair.Key + ").",
                            ex);
                    }
                }

                Outlook.ExplorerEvents_10_SelectionChangeEventHandler selectionChangeHandler;
                if (_explorerSelectionChangeHandlers.TryGetValue(pair.Key, out selectionChangeHandler))
                {
                    try
                    {
                        pair.Value.SelectionChange -= selectionChangeHandler;
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(
                            LogCategories.Core,
                            "Failed to unhook Explorer.SelectionChange (explorerKey=" + pair.Key + ").",
                            ex);
                    }
                }
            }
            _hookedExplorerEvents.Clear();
            _inlineResponseHandlers.Clear();
            _inlineResponseCloseHandlers.Clear();
            _explorerSelectionChangeHandlers.Clear();
            _explorerFolderSwitchHandlers.Clear();
            _inlineResponseSubscriptions.Clear();

            foreach (var pair in _hookedExplorers)
            {
                ComInteropScope.TryRelease(pair.Value, LogCategories.Core, "Failed to release tracked Explorer COM object.");
            }
            _hookedExplorers.Clear();

            if (_explorersEvents != null)
            {
                try
                {
                    _explorersEvents.NewExplorer -= OnNewExplorer;
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to unhook Explorers.NewExplorer.", ex);
                }
                _explorersEvents = null;
            }

            ComInteropScope.TryFinalRelease(_explorers, LogCategories.Core, "Failed to release Explorers COM object.");
            _explorers = null;
        }

        private void UnhookInspector()
        {
            if (_inspectors == null)
            {
                return;
            }
            try
            {
                _inspectors.NewInspector -= OnNewInspector;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to unhook Inspectors.NewInspector.", ex);
            }
            finally
            {
                ComInteropScope.TryFinalRelease(_inspectors, LogCategories.Core, "Failed to release Inspectors COM object.");
                _inspectors = null;
            }
        }

        private void OnNewInspector(Outlook.Inspector inspector)
        {
            if (inspector == null)
            {
                return;
            }
            try
            {
                var appointment = inspector.CurrentItem as Outlook.AppointmentItem;
                if (appointment != null)
                {
                    EnsureSubscriptionForAppointment(appointment);
                }
                var mail = inspector.CurrentItem as Outlook.MailItem;
                if (mail != null && IsMailComposeCandidate(mail, "new_inspector"))
                {
                    string inspectorIdentityKey = ComInteropScope.ResolveIdentityKey(inspector, LogCategories.FileLink, "Inspector");
                    MailComposeSubscription subscription = EnsureMailComposeSubscription(
                        mail,
                        inspectorIdentityKey,
                        false,
                        null,
                        inspector);
                    if (subscription != null && DiagnosticsLogger.IsEnabled)
                    {
                        LogFileLink(
                            "Compose subscription associated with Inspector surface (inspectorKey="
                            + (inspectorIdentityKey ?? string.Empty)
                            + ").");
                    }
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to process NewInspector event.", ex);
            }
        }

        private void OnNewExplorer(Outlook.Explorer explorer)
        {
            try
            {
                HookExplorer(explorer);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to hook new Explorer lifecycle.", ex);
            }
        }

        private void OnExplorerInlineResponse(string explorerKey, object item)
        {
            var mail = item as Outlook.MailItem;
            if (mail == null || !IsMailComposeCandidate(mail, "inline_response"))
            {
                return;
            }
            try
            {
                MailComposeSubscription subscription = EnsureMailComposeSubscription(mail, string.Empty, true, explorerKey);
                if (subscription == null)
                {
                    return;
                }

                MailComposeSubscription previousSubscription;
                if (_inlineResponseSubscriptions.TryGetValue(explorerKey, out previousSubscription)
                    && !ReferenceEquals(previousSubscription, subscription))
                {
                    previousSubscription.MarkInlineResponseClosed(explorerKey);
                }
                _inlineResponseSubscriptions[explorerKey] = subscription;
                if (DiagnosticsLogger.IsEnabled)
                {
                    LogFileLink(
                        "Compose subscription bound via Explorer.InlineResponse (explorerKey="
                        + (explorerKey ?? string.Empty)
                        + ").");
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Failed to bind compose subscription on Explorer.InlineResponse (explorerKey="
                    + (explorerKey ?? string.Empty)
                    + ").",
                    ex);
            }
        }

        private void OnExplorerInlineResponseClose(string explorerKey)
        {
            try
            {
                MailComposeSubscription subscription;
                if (!_inlineResponseSubscriptions.TryGetValue(explorerKey, out subscription))
                {
                    if (DiagnosticsLogger.IsEnabled)
                    {
                        LogFileLink(
                            "Explorer.InlineResponseClose ignored without tracked compose subscription (explorerKey="
                            + (explorerKey ?? string.Empty)
                            + ").");
                    }
                    return;
                }

                _inlineResponseSubscriptions.Remove(explorerKey);
                subscription.MarkInlineResponseClosed(explorerKey);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Failed to close inline compose surface (explorerKey="
                    + (explorerKey ?? string.Empty)
                    + ").",
                    ex);
            }
        }

        private static bool IsMailComposeCandidate(Outlook.MailItem mail, string reason)
        {
            if (mail == null)
            {
                return false;
            }
            try
            {
                if (mail.Sent)
                {
                    if (DiagnosticsLogger.IsEnabled)
                    {
                        LogFileLink("Mail compose subscription skipped (reason=" + (reason ?? "n/a") + ", sent=True).");
                    }
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.FileLink, "Failed to verify mail compose state (reason=" + (reason ?? "n/a") + ").", ex);
                return false;
            }
        }
    }
}

