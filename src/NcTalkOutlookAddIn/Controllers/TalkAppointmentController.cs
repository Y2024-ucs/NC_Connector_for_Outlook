// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Globalization;
using System.Net;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Controllers
{
    internal sealed partial class TalkAppointmentController
    {
        private readonly NextcloudTalkAddIn _owner;

        internal TalkAppointmentController(NextcloudTalkAddIn owner)
        {
            _owner = owner;
        }

        internal void ApplyRoomToAppointment(Outlook.AppointmentItem appointment, TalkRoomRequest request, TalkRoomCreationResult result)
        {
            // COM lifecycle guard: appointment/result may be unavailable during event teardown.
            if (appointment == null || result == null)
            {
                return;
            }

            NextcloudTalkAddIn.LogTalkMessage(
                "Writing talk data to appointment (token="
                + result.RoomToken
                + ", type="
                + request.RoomType
                + ", lobby="
                + request.LobbyEnabled
                + ", search="
                + request.SearchVisible
                + ", passwordSet="
                + (!string.IsNullOrEmpty(request.Password))
                + ", addUsers="
                + request.AddUsers
                + ", addGuests="
                + request.AddGuests
                + ", delegate="
                + (string.IsNullOrEmpty(request.DelegateModeratorId) ? "n/a" : request.DelegateModeratorId)
                + ").");

            if (!string.IsNullOrWhiteSpace(request.Title))
            {
                appointment.Subject = request.Title.Trim();
            }

            appointment.Location = result.RoomUrl;
            string normalizedDescriptionType = NextcloudTalkAddIn.NormalizeTalkEventDescriptionType(request.DescriptionType);
            if (string.Equals(normalizedDescriptionType, "html", StringComparison.OrdinalIgnoreCase))
            {
                // AppointmentItem HTML body read is intentionally disabled for compatibility.
                string updatedHtmlBody = TalkDescriptionTemplateController.UpdateHtmlBodyWithTalkBlock(
                    string.Empty,
                    appointment.Body,
                    result.RoomUrl,
                    request.Password,
                    request.DescriptionLanguage,
                    request.InvitationTemplate);
                if (!NextcloudTalkAddIn.TryWriteAppointmentHtmlBody(appointment, updatedHtmlBody))
                {
                    NextcloudTalkAddIn.LogTalkMessage("Appointment HTML editor unavailable, falling back to plain-text Talk block insertion.");
                    appointment.Body = TalkDescriptionTemplateController.UpdateBodyWithTalkBlock(
                        appointment.Body,
                        result.RoomUrl,
                        request.Password,
                        request.DescriptionLanguage,
                        request.InvitationTemplate);
                }
            }
            else
            {
                appointment.Body = TalkDescriptionTemplateController.UpdateBodyWithTalkBlock(
                    appointment.Body,
                    result.RoomUrl,
                    request.Password,
                    request.DescriptionLanguage,
                    request.InvitationTemplate);
            }

            if (!string.IsNullOrWhiteSpace(request.DelegateModeratorId))
            {
                string delegateId = request.DelegateModeratorId.Trim();
                string delegateName = !string.IsNullOrWhiteSpace(request.DelegateModeratorName) ? request.DelegateModeratorName.Trim() : delegateId;

                SetUserProperty(appointment, NextcloudTalkAddIn.IcalDelegate, Outlook.OlUserPropertyType.olText, delegateId);
                SetUserProperty(appointment, NextcloudTalkAddIn.IcalDelegateName, Outlook.OlUserPropertyType.olText, delegateName);
                SetUserProperty(appointment, NextcloudTalkAddIn.IcalDelegated, Outlook.OlUserPropertyType.olText, "FALSE");
                SetUserProperty(appointment, NextcloudTalkAddIn.IcalDelegateReady, Outlook.OlUserPropertyType.olText, "TRUE");
            }
            else
            {
                RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalDelegate);
                RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalDelegateName);
                RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalDelegated);
                RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalDelegateReady);
            }

            long ignoredStartEpoch;
            PersistCoreIcalProperties(
                appointment,
                result.RoomToken,
                result.RoomUrl,
                request.LobbyEnabled,
                result.CreatedAsEventConversation,
                out ignoredStartEpoch);
            SetUserProperty(appointment, NextcloudTalkAddIn.IcalAddUsers, Outlook.OlUserPropertyType.olText, request.AddUsers ? "TRUE" : "FALSE");
            SetUserProperty(appointment, NextcloudTalkAddIn.IcalAddGuests, Outlook.OlUserPropertyType.olText, request.AddGuests ? "TRUE" : "FALSE");
            NextcloudTalkAddIn.LogTalkMessage("X-NCTALK fields updated (token set, lobby=" + request.LobbyEnabled + ", search=" + request.SearchVisible + ").");

            _owner.RegisterSubscription(appointment, result);
            TalkAppointmentSyncSnapshot descriptionSnapshot =
                CaptureRemoteSyncSnapshot(
                    appointment,
                    result.RoomToken,
                    result.RoomUrl,
                    request.LobbyEnabled,
                    result.CreatedAsEventConversation,
                    null,
                    false);
            if (descriptionSnapshot != null)
            {
                descriptionSnapshot.RoomName = string.Empty;
                descriptionSnapshot.UpdateLobby = false;
                descriptionSnapshot.AddUsers = false;
                descriptionSnapshot.AddGuests = false;
                descriptionSnapshot.AttendeeEmails.Clear();
                descriptionSnapshot.DelegationPending = false;
                descriptionSnapshot.DelegateUserId = string.Empty;
                _owner.QueueTalkAppointmentSync(
                    descriptionSnapshot);
            }
        }

        internal bool PersistCoreIcalProperties(
            Outlook.AppointmentItem appointment,
            string roomToken,
            string roomUrl,
            bool lobbyEnabled,
            bool isEventConversation,
            out long startEpoch)
        {
            startEpoch = 0;
            if (appointment == null || string.IsNullOrWhiteSpace(roomToken))
            {
                return false;
            }

            string normalizedRoomToken = roomToken.Trim();
            SetUserProperty(appointment, NextcloudTalkAddIn.IcalToken, Outlook.OlUserPropertyType.olText, normalizedRoomToken);
            if (!string.IsNullOrWhiteSpace(roomUrl))
            {
                SetUserProperty(appointment, NextcloudTalkAddIn.IcalUrl, Outlook.OlUserPropertyType.olText, roomUrl.Trim());
            }
            SetUserProperty(appointment, NextcloudTalkAddIn.IcalLobby, Outlook.OlUserPropertyType.olText, lobbyEnabled ? "TRUE" : "FALSE");
            SetUserProperty(appointment, NextcloudTalkAddIn.IcalEvent, Outlook.OlUserPropertyType.olText, isEventConversation ? "event" : "standard");

            DateTime start;
            DateTime end;
            try
            {
                start = appointment.Start;
                end = appointment.End;
            }
            catch (Exception ex)
            {
                RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalStart);
                RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalObjectId);
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to read appointment time while persisting X-NCTALK core fields (token=" + normalizedRoomToken + ").", ex);
                return false;
            }

            long? startValue = TimeUtilities.ToUnixTimeSeconds(start);
            long? endValue = TimeUtilities.ToUnixTimeSeconds(end);
            if (!startValue.HasValue || startValue.Value <= 0)
            {
                RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalStart);
                RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalObjectId);
                NextcloudTalkAddIn.LogTalkMessage("X-NCTALK core fields persisted without valid start (token=" + normalizedRoomToken + ", lobby=" + lobbyEnabled + ", event=" + isEventConversation + ").");
                return false;
            }

            startEpoch = startValue.Value;
            SetUserProperty(appointment, NextcloudTalkAddIn.IcalStart, Outlook.OlUserPropertyType.olText, startEpoch.ToString(CultureInfo.InvariantCulture));
            if (endValue.HasValue && endValue.Value > 0)
            {
                string objectId = startEpoch.ToString(CultureInfo.InvariantCulture) + "#" + endValue.Value.ToString(CultureInfo.InvariantCulture);
                SetUserProperty(appointment, NextcloudTalkAddIn.IcalObjectId, Outlook.OlUserPropertyType.olText, objectId);
            }
            else
            {
                RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalObjectId);
            }

            NextcloudTalkAddIn.LogTalkMessage("X-NCTALK core fields persisted (token=" + normalizedRoomToken + ", lobby=" + lobbyEnabled + ", event=" + isEventConversation + ", startEpoch=" + startEpoch.ToString(CultureInfo.InvariantCulture) + ").");
            return true;
        }

        internal static long? GetIcalStartEpochOrNull(Outlook.AppointmentItem appointment)
        {
            string rawValue = GetUserPropertyText(appointment, NextcloudTalkAddIn.IcalStart);
            if (string.IsNullOrWhiteSpace(rawValue))
            {
                return null;
            }

            long parsed;
            if (!long.TryParse(rawValue.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed) || parsed <= 0)
            {
                return null;
            }
            return parsed;
        }

        internal void ResolveRuntimeRoomTraits(
            Outlook.AppointmentItem appointment,
            string roomToken,
            bool fallbackLobbyEnabled,
            bool fallbackIsEventConversation,
            out bool lobbyKnown,
            out bool lobbyEnabled,
            out bool isEventConversation)
        {
            bool? resolvedLobby = null;
            bool? resolvedEventConversation = null;

            try
            {
                bool hasLobbyFlag = HasUserProperty(appointment, NextcloudTalkAddIn.IcalLobby);
                if (hasLobbyFlag)
                {
                    resolvedLobby = GetUserPropertyBool(appointment, NextcloudTalkAddIn.IcalLobby);
                }

                TalkRoomType? roomType = GetRoomType(appointment);
                if (roomType.HasValue)
                {
                    resolvedEventConversation = roomType.Value == TalkRoomType.EventConversation;
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to resolve runtime room traits (token=" + (roomToken ?? "n/a") + ").", ex);
            }

            lobbyKnown = resolvedLobby.HasValue;
            lobbyEnabled = resolvedLobby.HasValue ? resolvedLobby.Value : fallbackLobbyEnabled;
            isEventConversation = resolvedEventConversation.HasValue ? resolvedEventConversation.Value : fallbackIsEventConversation;
        }

        internal bool IsDelegatedToOtherUser(Outlook.AppointmentItem appointment, out string delegateId)
        {
            delegateId = GetUserPropertyText(appointment, NextcloudTalkAddIn.IcalDelegate) ?? string.Empty;
            bool delegated = GetUserPropertyBool(appointment, NextcloudTalkAddIn.IcalDelegated);

            if (!delegated || string.IsNullOrWhiteSpace(delegateId))
            {
                delegateId = string.Empty;
                return false;
            }
            string currentUser = _owner.CurrentSettings != null ? (_owner.CurrentSettings.Username ?? string.Empty) : string.Empty;
            if (string.IsNullOrWhiteSpace(currentUser))
            {
                return true;
            }
            return !string.Equals(delegateId.Trim(), currentUser.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        internal bool IsDelegationPending(Outlook.AppointmentItem appointment, out string delegateId)
        {
            delegateId = GetUserPropertyText(appointment, NextcloudTalkAddIn.IcalDelegate) ?? string.Empty;
            bool delegated = GetUserPropertyBool(appointment, NextcloudTalkAddIn.IcalDelegated);

            if (delegated)
            {
                delegateId = string.Empty;
                return false;
            }
            return !string.IsNullOrWhiteSpace(delegateId);
        }

        private static string BuildDescriptionPayload(Outlook.AppointmentItem appointment)
        {
            // Defensive fallback: upstream callers are internal, but COM callbacks can still race.
            if (appointment == null)
            {
                return null;
            }
            string body = appointment.Body ?? string.Empty;
            return body.Trim();
        }

        internal static TalkRoomType? GetRoomType(Outlook.AppointmentItem appointment)
        {
            string eventRaw = GetUserPropertyText(appointment, NextcloudTalkAddIn.IcalEvent);
            if (!string.IsNullOrWhiteSpace(eventRaw))
            {
                string normalized = eventRaw.Trim();
                if (string.Equals(normalized, "event", StringComparison.OrdinalIgnoreCase))
                {
                    return TalkRoomType.EventConversation;
                }
                if (string.Equals(normalized, "standard", StringComparison.OrdinalIgnoreCase))
                {
                    return TalkRoomType.StandardRoom;
                }
            }
            return null;
        }

        private void PersistEventConversationTraits(Outlook.AppointmentItem appointment, string roomToken)
        {
            if (appointment == null)
            {
                return;
            }
            try
            {
                SetUserProperty(appointment, NextcloudTalkAddIn.IcalEvent, Outlook.OlUserPropertyType.olText, "event");
                NextcloudTalkAddIn.LogTalkMessage("Conversation type persisted as event (token=" + (roomToken ?? "n/a") + ").");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to persist event conversation traits (token=" + (roomToken ?? "n/a") + ").", ex);
            }
        }

        internal void ClearTalkProperties(Outlook.AppointmentItem appointment)
        {
            RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalToken);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalUrl);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalLobby);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalStart);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalEvent);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalObjectId);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalAddUsers);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalAddGuests);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalDelegate);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalDelegateName);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalDelegated);
            RemoveUserProperty(appointment, NextcloudTalkAddIn.IcalDelegateReady);
        }

        private static bool IsEventConversationDescriptionError(TalkServiceException ex)
        {
            if (ex == null || string.IsNullOrWhiteSpace(ex.Message))
            {
                return false;
            }
            string normalized = ex.Message.Trim();
            return string.Equals(normalized, "event", StringComparison.OrdinalIgnoreCase)
                   || normalized.IndexOf("event conversation", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool IsMissingOrForbiddenRoomMutationError(TalkServiceException ex)
        {
            if (ex == null)
            {
                return false;
            }
            return ex.StatusCode == HttpStatusCode.NotFound || ex.StatusCode == HttpStatusCode.Forbidden;
        }

        private static string GetNormalizedRoomName(Outlook.AppointmentItem appointment)
        {
            // Keep null-tolerant: caller paths run inside Outlook event handlers with COM races.
            if (appointment == null)
            {
                return string.Empty;
            }
            try
            {
                string subject = appointment.Subject;
                return string.IsNullOrWhiteSpace(subject) ? string.Empty : subject.Trim();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Talk, "Failed to read appointment subject for room name sync.", ex);
                return string.Empty;
            }
        }
    }
}

