// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
using System.Runtime.InteropServices;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
    public sealed partial class NextcloudTalkAddIn
    {
        private sealed class AppointmentSubscription : IDisposable
        {
            private readonly NextcloudTalkAddIn _owner;
            private readonly Outlook.AppointmentItem _appointment;
            private readonly string _key;
            private readonly string _roomToken;
            private readonly string _roomUrl;
            private readonly bool _lobbyEnabled;
            private readonly Outlook.ItemEvents_10_Event _events;
            private readonly bool _isEventConversation;
            private long? _lastLobbyTimer;
            private bool _roomDeleted;
            private bool _disposed;
            private bool _unsavedCloseCleanupPending;
            private int _unsavedCloseCleanupAttempts;
            private System.Windows.Forms.Timer _unsavedCloseCleanupTimer;
            private System.Windows.Forms.Timer _deferredWriteLobbyTimer;
            private int _deferredWriteLobbyAttempts;
            private string _entryId;
            private const int UnsavedCloseCleanupMaxAttempts = 90;
            private const int DeferredWriteLobbyMaxAttempts = 4;

            internal AppointmentSubscription(
                NextcloudTalkAddIn owner,
                Outlook.AppointmentItem appointment,
                string key,
                string roomToken,
                string roomUrl,
                bool lobbyEnabled,
                bool isEventConversation,
                string entryId)
            {
                _owner = owner;
                _appointment = appointment;
                _key = key;
                _roomToken = roomToken;
                _roomUrl = roomUrl;
                _lobbyEnabled = lobbyEnabled;
                _isEventConversation = isEventConversation;
                _lastLobbyTimer = GetIcalStartEpochOrNull(appointment);
                _entryId = entryId;
                _events = appointment as Outlook.ItemEvents_10_Event;
                AttachEventHandlers();

                LogTalk("Subscription registered (token=" + _roomToken + ", lobby=" + _lobbyEnabled + ", event=" + _isEventConversation + ", EntryId=" + (_entryId ?? "n/a") + ").");
            }

            private void AttachEventHandlers()
            {
                if (_events == null)
                {
                    return;
                }

                bool beforeDeleteAttached = false;
                bool writeAttached = false;
                bool closeAttached = false;
                try
                {
                    _events.BeforeDelete += OnBeforeDelete;
                    beforeDeleteAttached = true;
                    _events.Write += OnWrite;
                    writeAttached = true;
                    _events.Close += OnClose;
                    closeAttached = true;
                }
                catch
                {
                    DetachEventHandlers(
                        beforeDeleteAttached,
                        writeAttached,
                        closeAttached);
                    throw;
                }
            }

            private void DetachEventHandlers(
                bool beforeDeleteAttached,
                bool writeAttached,
                bool closeAttached)
            {
                if (_events == null)
                {
                    return;
                }

                if (closeAttached)
                {
                    try
                    {
                        _events.Close -= OnClose;
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to detach the Talk appointment Close event.", ex);
                    }
                }
                if (writeAttached)
                {
                    try
                    {
                        _events.Write -= OnWrite;
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to detach the Talk appointment Write event.", ex);
                    }
                }
                if (beforeDeleteAttached)
                {
                    try
                    {
                        _events.BeforeDelete -= OnBeforeDelete;
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to detach the Talk appointment BeforeDelete event.", ex);
                    }
                }
            }

            private void OnWrite(ref bool cancel)
            {
                if (_unsavedCloseCleanupPending)
                {
                    _unsavedCloseCleanupPending = false;
                    _unsavedCloseCleanupAttempts = 0;
                    StopUnsavedCloseCleanupTimer();
                    LogTalk("OnWrite canceled pending unsaved cleanup (token=" + _roomToken + ").");
                }
                if (string.IsNullOrWhiteSpace(_roomToken))
                {
                    LogTalk("OnWrite ignored (no token).");
                    return;
                }

                if (!_owner.IsOrganizer(_appointment))
                {
                    LogTalk("OnWrite ignored (not organizer, token=" + _roomToken + ").");
                    _owner.RefreshEntryBinding(this);
                    return;
                }
                string delegateId;
                if (_owner._talkAppointmentController.IsDelegatedToOtherUser(_appointment, out delegateId))
                {
                    LogTalk("OnWrite skipped (delegation=" + delegateId + ", token=" + _roomToken + ").");
                    _owner.RefreshEntryBinding(this);
                    return;
                }

                LogTalk("OnWrite for appointment (token=" + _roomToken + ").");
                TalkAppointmentSyncSnapshot snapshot =
                    _owner._talkAppointmentController.CaptureRemoteSyncSnapshot(
                    _appointment,
                    _roomToken,
                    _roomUrl,
                    _lobbyEnabled,
                    _isEventConversation,
                    _lastLobbyTimer,
                    false);
                if (snapshot != null)
                {
                    _owner.QueueTalkAppointmentSync(snapshot);
                    LogTalk(
                        "OnWrite queued background Talk synchronization (token="
                        + _roomToken
                        + ", lobby="
                        + snapshot.UpdateLobby
                        + ", participants="
                        + snapshot.AttendeeEmails.Count
                        + ", delegation="
                        + snapshot.DelegationPending
                        + ").");
                    ScheduleDeferredWriteLobbyVerification();
                }
                _owner.RefreshEntryBinding(this);
            }

            private void OnBeforeDelete(object item, ref bool cancel)
            {
                if (cancel)
                {
                    LogTalk("BeforeDelete ignored because deletion was already cancelled (token=" + _roomToken + ").");
                    return;
                }
                if (string.IsNullOrWhiteSpace(_roomToken))
                {
                    LogTalk("BeforeDelete ignored (no token).");
                    return;
                }
                // An attendee can delete only a calendar copy. The organizer owns the room
                // shared by the appointment's other participants.
                if (!_owner.IsOrganizer(_appointment))
                {
                    LogTalk("BeforeDelete ignored (not organizer, token=" + _roomToken + ").");
                    return;
                }
                if (!IsRoomDeletionAllowedForRecurrence())
                {
                    Dispose();
                    return;
                }

                QueueSavedTalkRoomDeletion();
            }

            private void OnClose(ref bool cancel)
            {
                if (!_owner.IsOrganizer(_appointment))
                {
                    LogTalk("OnClose without organizer (token=" + _roomToken + ").");
                    Dispose();
                    return;
                }
                if (!_roomDeleted && _appointment != null && !_appointment.Saved)
                {
                    LogTalk("OnClose unsaved state observed (token=" + _roomToken + ", cancel=" + cancel + ").");
                    if (!cancel)
                    {
                        ScheduleUnsavedCloseCleanup();
                    }
                    return;
                }

                LogTalk("OnClose completed (token=" + _roomToken + ", deleted=" + _roomDeleted + ").");
            }

            private bool IsRoomDeletionAllowedForRecurrence()
            {
                // An occurrence or exception shares its series room. Removing one must not
                // tear down the room used by the remaining series.
                try
                {
                    Outlook.OlRecurrenceState recurrenceState =
                        _appointment.RecurrenceState;
                    if (recurrenceState == Outlook.OlRecurrenceState.olApptNotRecurring
                        || recurrenceState == Outlook.OlRecurrenceState.olApptMaster)
                    {
                        return true;
                    }

                    LogTalk(
                        "BeforeDelete retained the room for a recurring occurrence or exception (token="
                        + _roomToken
                        + ", recurrenceState="
                        + recurrenceState
                        + ").");
                    return false;
                }
                catch (COMException ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Talk,
                        "BeforeDelete retained the room because Outlook did not expose the recurrence state.",
                        ex);
                    return false;
                }
            }

            private void ScheduleUnsavedCloseCleanup()
            {
                if (_unsavedCloseCleanupPending)
                {
                    return;
                }

                _unsavedCloseCleanupPending = true;
                _unsavedCloseCleanupAttempts = 0;
                _unsavedCloseCleanupTimer = new System.Windows.Forms.Timer();
                _unsavedCloseCleanupTimer.Interval = 1000;
                _unsavedCloseCleanupTimer.Tick += OnUnsavedCloseCleanupTick;
                _unsavedCloseCleanupTimer.Start();
                LogTalk("OnClose unsaved cleanup scheduled (token=" + _roomToken + ").");
            }

            private void ScheduleDeferredWriteLobbyVerification()
            {
                if (_disposed || _roomDeleted || _appointment == null)
                {
                    return;
                }

                _deferredWriteLobbyAttempts = 0;
                if (_deferredWriteLobbyTimer == null)
                {
                    _deferredWriteLobbyTimer = new System.Windows.Forms.Timer();
                    _deferredWriteLobbyTimer.Interval = 750;
                    _deferredWriteLobbyTimer.Tick += OnDeferredWriteLobbyTick;
                }

                _deferredWriteLobbyTimer.Stop();
                _deferredWriteLobbyTimer.Start();
                LogTalk("Deferred post-write lobby verification scheduled (token=" + _roomToken + ").");
            }

            private void OnDeferredWriteLobbyTick(object sender, EventArgs e)
            {
                try
                {
                    OnDeferredWriteLobbyTickCore();
                }
                catch (COMException ex)
                {
                    StopDeferredWriteLobbyTimerAfterError();
                    DiagnosticsLogger.LogException(LogCategories.Talk, "Deferred post-write lobby verification stopped after Outlook invalidated the appointment.", ex);
                }
                catch (Exception ex)
                {
                    StopDeferredWriteLobbyTimerAfterError();
                    DiagnosticsLogger.LogException(LogCategories.Talk, "Deferred post-write lobby verification failed.", ex);
                }
            }

            private void OnDeferredWriteLobbyTickCore()
            {
                if (_disposed || _roomDeleted || _appointment == null)
                {
                    _deferredWriteLobbyAttempts = 0;
                    StopDeferredWriteLobbyTimer();
                    return;
                }

                _deferredWriteLobbyAttempts++;
                _owner.RefreshEntryBinding(this);

                if (!_owner.IsOrganizer(_appointment))
                {
                    _deferredWriteLobbyAttempts = 0;
                    StopDeferredWriteLobbyTimer();
                    LogTalk("Deferred post-write lobby verification skipped (not organizer, token=" + _roomToken + ").");
                    return;
                }
                string delegateId;
                if (_owner._talkAppointmentController.IsDelegatedToOtherUser(_appointment, out delegateId))
                {
                    _deferredWriteLobbyAttempts = 0;
                    StopDeferredWriteLobbyTimer();
                    LogTalk("Deferred post-write lobby verification skipped (delegation=" + delegateId + ", token=" + _roomToken + ").");
                    return;
                }
                TalkAppointmentSyncSnapshot snapshot =
                    _owner._talkAppointmentController.CaptureRemoteSyncSnapshot(
                        _appointment,
                        _roomToken,
                        _roomUrl,
                        _lobbyEnabled,
                        _isEventConversation,
                        _lastLobbyTimer,
                        true);
                if (snapshot == null
                    || (!snapshot.LobbyEnabled && snapshot.LobbyKnown))
                {
                    _deferredWriteLobbyAttempts = 0;
                    StopDeferredWriteLobbyTimer();
                    return;
                }
                if (!snapshot.UpdateLobby)
                {
                    if (_deferredWriteLobbyAttempts >= DeferredWriteLobbyMaxAttempts)
                    {
                        _deferredWriteLobbyAttempts = 0;
                        StopDeferredWriteLobbyTimer();
                        LogTalk(
                            "Deferred post-write lobby verification completed without a new start value (token="
                            + _roomToken
                            + ").");
                    }
                    return;
                }

                _owner.QueueTalkAppointmentSync(snapshot);
                _deferredWriteLobbyAttempts = 0;
                StopDeferredWriteLobbyTimer();
                LogTalk(
                    "Deferred post-write lobby update queued (token="
                    + _roomToken
                    + ", startEpoch="
                    + snapshot.StartEpoch.ToString(CultureInfo.InvariantCulture)
                    + ").");
            }

            private void StopDeferredWriteLobbyTimerAfterError()
            {
                _deferredWriteLobbyAttempts = 0;
                try
                {
                    StopDeferredWriteLobbyTimer();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to stop deferred lobby verification timer after error.", ex);
                }
            }

            private void OnUnsavedCloseCleanupTick(object sender, EventArgs e)
            {
                if (_disposed || _roomDeleted || _appointment == null)
                {
                    _unsavedCloseCleanupPending = false;
                    _unsavedCloseCleanupAttempts = 0;
                    StopUnsavedCloseCleanupTimer();
                    return;
                }

                _unsavedCloseCleanupAttempts++;
                bool saved;
                bool hasSavedState = TryGetAppointmentSaved(_appointment, out saved);
                bool inspectorOpen = IsAppointmentOpenInAnyInspector(_appointment);
                if (saved)
                {
                    _unsavedCloseCleanupPending = false;
                    _unsavedCloseCleanupAttempts = 0;
                    StopUnsavedCloseCleanupTimer();
                    LogTalk("Deferred OnClose cleanup skipped (token=" + _roomToken + ", saved=True).");
                    return;
                }
                if (inspectorOpen)
                {
                    if (_unsavedCloseCleanupAttempts >= UnsavedCloseCleanupMaxAttempts)
                    {
                        _unsavedCloseCleanupPending = false;
                        _unsavedCloseCleanupAttempts = 0;
                        StopUnsavedCloseCleanupTimer();
                        LogTalk("Deferred OnClose cleanup skipped after timeout (token=" + _roomToken + ", saved=" + saved + ", hasSavedState=" + hasSavedState + ", inspectorOpen=True).");
                        return;
                    }
                    if (_unsavedCloseCleanupAttempts == 1 || (_unsavedCloseCleanupAttempts % 5) == 0)
                    {
                        LogTalk("Deferred OnClose cleanup waiting for inspector close (token=" + _roomToken + ", attempt=" + _unsavedCloseCleanupAttempts.ToString(CultureInfo.InvariantCulture) + ", saved=" + saved + ", hasSavedState=" + hasSavedState + ").");
                    }
                    return;
                }

                _unsavedCloseCleanupPending = false;
                _unsavedCloseCleanupAttempts = 0;
                StopUnsavedCloseCleanupTimer();

                if (!saved)
                {
                    LogTalk("Deferred OnClose cleanup deleting room (token=" + _roomToken + ").");
                    EnsureRoomDeleted();
                    return;
                }

                LogTalk("Deferred OnClose cleanup skipped (token=" + _roomToken + ", saved=True, hasSavedState=" + hasSavedState + ", inspectorOpen=False).");
            }

            private static bool TryGetAppointmentSaved(Outlook.AppointmentItem appointment, out bool saved)
            {
                saved = false;
                if (appointment == null)
                {
                    return false;
                }
                try
                {
                    saved = appointment.Saved;
                    return true;
                }
                catch
                {
                    return false;
                }
            }

            private void StopUnsavedCloseCleanupTimer()
            {
                if (_unsavedCloseCleanupTimer == null)
                {
                    return;
                }

                _unsavedCloseCleanupTimer.Stop();
                _unsavedCloseCleanupTimer.Tick -= OnUnsavedCloseCleanupTick;
                _unsavedCloseCleanupTimer.Dispose();
                _unsavedCloseCleanupTimer = null;
            }

            private void StopDeferredWriteLobbyTimer()
            {
                if (_deferredWriteLobbyTimer == null)
                {
                    return;
                }

                _deferredWriteLobbyTimer.Stop();
                _deferredWriteLobbyTimer.Tick -= OnDeferredWriteLobbyTick;
                _deferredWriteLobbyTimer.Dispose();
                _deferredWriteLobbyTimer = null;
            }

            private bool IsAppointmentOpenInAnyInspector(Outlook.AppointmentItem appointment)
            {
                if (appointment == null || _owner == null || _owner._inspectors == null)
                {
                    return false;
                }
                string appointmentEntryId = null;
                try
                {
                    appointmentEntryId = appointment.EntryID;
                }
                catch
                {
                }
                int inspectorCount = 0;
                try
                {
                    inspectorCount = _owner._inspectors.Count;
                }
                catch
                {
                    return false;
                }
                for (int i = 1; i <= inspectorCount; i++)
                {
                    Outlook.Inspector inspector = null;
                    object currentItem = null;
                    Outlook.AppointmentItem currentAppointment = null;
                    try
                    {
                        inspector = _owner._inspectors[i];
                        if (inspector == null)
                        {
                            continue;
                        }

                        currentItem = inspector.CurrentItem;
                        currentAppointment = currentItem as Outlook.AppointmentItem;
                        if (currentAppointment == null)
                        {
                            continue;
                        }
                        if (currentAppointment == appointment)
                        {
                            return true;
                        }
                        if (!string.IsNullOrWhiteSpace(appointmentEntryId))
                        {
                            string currentEntryId = null;
                            try
                            {
                                currentEntryId = currentAppointment.EntryID;
                            }
                            catch
                            {
                            }
                            if (!string.IsNullOrWhiteSpace(currentEntryId)
                                && string.Equals(currentEntryId, appointmentEntryId, StringComparison.OrdinalIgnoreCase))
                            {
                                return true;
                            }
                        }
                    }
                    catch
                    {
                    }
                    finally
                    {
                        if (currentAppointment != null && !ReferenceEquals(currentAppointment, appointment))
                        {
                            ComInteropScope.TryRelease(
                                currentAppointment,
                                LogCategories.Talk,
                                "Failed to release current appointment COM object during unsaved-close lookup.");
                        }
                        if (currentItem != null
                            && !ReferenceEquals(currentItem, currentAppointment)
                            && !ReferenceEquals(currentItem, appointment))
                        {
                            ComInteropScope.TryRelease(
                                currentItem,
                                LogCategories.Talk,
                                "Failed to release current item COM object during unsaved-close lookup.");
                        }
                        if (inspector != null)
                        {
                            ComInteropScope.TryRelease(
                                inspector,
                                LogCategories.Talk,
                                "Failed to release inspector COM object during unsaved-close lookup.");
                        }
                    }
                }
                return false;
            }

            private void EnsureRoomDeleted()
            {
                if (_roomDeleted || !_owner.IsOrganizer(_appointment))
                {
                    if (_roomDeleted)
                    {
                        LogTalk("EnsureRoomDeleted: room already deleted (token=" + _roomToken + ").");
                    }
                    return;
                }
                string delegateId;
                if (_owner._talkAppointmentController.IsDelegatedToOtherUser(_appointment, out delegateId))
                {
                    LogTalk("EnsureRoomDeleted skipped (delegation=" + delegateId + ", token=" + _roomToken + ").");
                    _owner._talkAppointmentController.ClearTalkProperties(_appointment);
                    _roomDeleted = true;
                    Dispose();
                    return;
                }
                bool queued = _owner.QueueUnsavedTalkRoomDeletion(
                    _roomToken,
                    _isEventConversation);
                _owner._talkAppointmentController.ClearTalkProperties(
                    _appointment);
                _roomDeleted = true;
                LogTalk(
                    (queued
                        ? "EnsureRoomDeleted queued persistent cleanup (token="
                        : "EnsureRoomDeleted could not queue persistent cleanup (token=")
                    + _roomToken
                    + ").");
                Dispose();
            }

            private void QueueSavedTalkRoomDeletion()
            {
                if (_roomDeleted || string.IsNullOrWhiteSpace(_roomToken))
                {
                    return;
                }

                string delegateId;
                if (_owner._talkAppointmentController.IsDelegatedToOtherUser(
                        _appointment,
                        out delegateId))
                {
                    LogTalk(
                        "Saved-event room deletion skipped (delegation="
                        + delegateId
                        + ", token="
                        + _roomToken
                        + ").");
                    _roomDeleted = true;
                    Dispose();
                    return;
                }

                _roomDeleted = true;
                bool queued = _owner.QueueSavedTalkRoomDeletion(
                    _roomToken,
                    _isEventConversation);
                LogTalk(
                    (queued
                        ? "BeforeDelete queued persistent room deletion (token="
                        : "BeforeDelete could not queue persistent room deletion (token=")
                    + _roomToken
                    + ").");
                Dispose();
            }

            public void Dispose()
            {
                if (_disposed)
                {
                    LogTalk("Subscription.Dispose called again (token=" + _roomToken + ").");
                    return;
                }
                _disposed = true;
                DetachEventHandlers(true, true, true);

                try
                {
                    StopUnsavedCloseCleanupTimer();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to stop the unsaved Talk appointment cleanup timer.", ex);
                }
                try
                {
                    StopDeferredWriteLobbyTimer();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to stop the deferred Talk appointment write timer.", ex);
                }

                try
                {
                    _owner.UnregisterSubscription(_key, _roomToken, _entryId);
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to unregister the Talk appointment subscription.", ex);
                }
                LogTalk("Subscription.Dispose completed (token=" + _roomToken + ").");
            }

            internal Outlook.AppointmentItem Appointment
            {
                get { return _appointment; }
            }

            internal string EntryId
            {
                get { return _entryId; }
            }

            internal bool IsFor(Outlook.AppointmentItem appointment)
            {
                if (appointment == null)
                {
                    return false;
                }
                if (appointment == _appointment)
                {
                    return true;
                }
                string thisEntryId = _entryId;
                if (string.IsNullOrWhiteSpace(thisEntryId))
                {
                    thisEntryId = GetEntryId(_appointment);
                }
                string otherEntryId = GetEntryId(appointment);
                if (!string.IsNullOrWhiteSpace(thisEntryId)
                    && !string.IsNullOrWhiteSpace(otherEntryId)
                    && string.Equals(thisEntryId, otherEntryId, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
                return false;
            }

            internal bool MatchesToken(string token)
            {
                return string.Equals(_roomToken, token, StringComparison.OrdinalIgnoreCase);
            }

            internal bool ApplyRemoteSyncResult(TalkAppointmentSyncResult result)
            {
                if (_disposed
                    || result == null
                    || !MatchesToken(result.RoomToken))
                {
                    return false;
                }
                if (result.AppliedLobbyEpoch.HasValue)
                {
                    _lastLobbyTimer = result.AppliedLobbyEpoch.Value;
                }
                return _owner._talkAppointmentController.ApplyRemoteSyncResult(
                    _appointment,
                    result);
            }

            internal void UpdateEntryId(string entryId)
            {
                _entryId = entryId;
                LogTalk("Subscription EntryId updated (token=" + _roomToken + ", EntryId=" + (_entryId ?? "n/a") + ").");
            }
        }

        private void DisposeTalkAppointmentSubscriptions()
        {
            var subscriptions =
                new AppointmentSubscription[
                    _activeSubscriptions.Count];
            _activeSubscriptions.Values.CopyTo(
                subscriptions,
                0);
            for (int i = 0; i < subscriptions.Length; i++)
            {
                AppointmentSubscription subscription =
                    subscriptions[i];
                if (subscription == null)
                {
                    continue;
                }
                try
                {
                    subscription.Dispose();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.Talk,
                        "Failed to dispose an appointment subscription during add-in teardown.",
                        ex);
                }
            }
            _activeSubscriptions.Clear();
            _subscriptionByToken.Clear();
            _subscriptionByEntryId.Clear();
        }

    }
}
