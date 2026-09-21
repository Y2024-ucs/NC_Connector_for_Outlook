// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.UI;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn
{
    public sealed partial class NextcloudTalkAddIn
    {
        internal sealed partial class MailComposeSubscription
        {
            private AttachmentAutomationSettings _attachmentAutomationSettingsSnapshot;
            private DateTime _attachmentAutomationSettingsSnapshotUtc;
            private Task<AttachmentAutomationSettings> _attachmentAutomationSettingsRefreshTask;
            private int _attachmentAutomationSettingsRefreshGeneration;
            private static readonly TimeSpan AttachmentAutomationSettingsCacheLifetime =
                TimeSpan.FromMinutes(5);

            private sealed class AttachmentAutomationSettings
            {
                internal bool AlwaysConnector { get; set; }

                internal bool OfferAboveEnabled { get; set; }

                internal int ThresholdMb { get; set; }

                internal long ThresholdBytes { get; set; }
            }

            private AttachmentAutomationSettings ReadAttachmentAutomationSettings()
            {
                if (HasFreshAttachmentAutomationSettingsSnapshot())
                {
                    return _attachmentAutomationSettingsSnapshot;
                }

                // Do not make sending mail depend on a live backend-policy refresh.
                // When the cached policy expires, refresh it in the background and
                // continue with the last known snapshot. This avoids a false
                // "attachment routing required" error every time the 5-minute
                // policy cache expires.
                AttachmentAutomationSettings staleSnapshot =
                    _attachmentAutomationSettingsSnapshot;
                BeginAttachmentAutomationSettingsRefresh();
                return staleSnapshot ?? ReadLocalAttachmentAutomationSettings();
            }

            private async Task<AttachmentAutomationSettings> ReadAttachmentAutomationSettingsAsync()
            {
                if (HasFreshAttachmentAutomationSettingsSnapshot())
                {
                    return _attachmentAutomationSettingsSnapshot;
                }

                while (!_disposed)
                {
                    BeginAttachmentAutomationSettingsRefresh();
                    Task<AttachmentAutomationSettings> refreshTask =
                        _attachmentAutomationSettingsRefreshTask;
                    if (refreshTask == null)
                    {
                        return ReadLocalAttachmentAutomationSettings();
                    }

                    AttachmentAutomationSettings refreshed =
                        await refreshTask;
                    if (ReferenceEquals(
                        refreshTask,
                        _attachmentAutomationSettingsRefreshTask))
                    {
                        return _attachmentAutomationSettingsSnapshot
                               ?? refreshed;
                    }
                }
                return ReadLocalAttachmentAutomationSettings();
            }

            private bool HasFreshAttachmentAutomationSettingsSnapshot()
            {
                return _attachmentAutomationSettingsSnapshot != null
                       && DateTime.UtcNow
                       - _attachmentAutomationSettingsSnapshotUtc
                       < AttachmentAutomationSettingsCacheLifetime;
            }

            internal void RefreshAttachmentAutomationSettings()
            {
                if (_disposed)
                {
                    return;
                }

                _attachmentAutomationSettingsRefreshGeneration++;
                _attachmentAutomationSettingsSnapshot = null;
                _attachmentAutomationSettingsSnapshotUtc =
                    DateTime.MinValue;
                _attachmentAutomationSettingsRefreshTask = null;
                BeginAttachmentAutomationSettingsRefresh();
                LogFileLink(
                    "Compose attachment settings refresh requested (composeKey="
                    + _composeKey
                    + ").");
            }

            private void BeginAttachmentAutomationSettingsRefresh()
            {
                if (_disposed
                    || (_attachmentAutomationSettingsRefreshTask != null
                        && !_attachmentAutomationSettingsRefreshTask.IsCompleted))
                {
                    return;
                }

                AttachmentAutomationSettings local =
                    ReadLocalAttachmentAutomationSettings();
                AddinSettings current = _owner._currentSettings;
                if (current == null)
                {
                    _attachmentAutomationSettingsSnapshot = local;
                    _attachmentAutomationSettingsSnapshotUtc = DateTime.UtcNow;
                    return;
                }

                var configuration = new TalkServiceConfiguration(
                    current.ServerUrl,
                    current.Username,
                    current.AppPassword);
                int refreshGeneration =
                    _attachmentAutomationSettingsRefreshGeneration;
                _attachmentAutomationSettingsRefreshTask = RefreshAttachmentAutomationSettingsAsync(
                    local,
                    configuration,
                    refreshGeneration);
                RunAttachmentFlowTask(
                    _attachmentAutomationSettingsRefreshTask,
                    "Compose attachment policy refresh failed");
            }

            private async Task<AttachmentAutomationSettings> RefreshAttachmentAutomationSettingsAsync(
                AttachmentAutomationSettings local,
                TalkServiceConfiguration configuration,
                int refreshGeneration)
            {
                AttachmentAutomationSettings resolved = local;
                if (configuration != null && configuration.IsComplete())
                {
                    BackendPolicyStatus policyStatus = await Task.Run(
                        () => _owner.FetchBackendPolicyStatus(
                            configuration,
                            "compose_attachment_evaluate")).ConfigureAwait(false);
                    resolved = ApplyAttachmentAutomationPolicy(
                        local,
                        policyStatus);
                }

                if (!_disposed
                    && refreshGeneration
                    == _attachmentAutomationSettingsRefreshGeneration)
                {
                    _attachmentAutomationSettingsSnapshot = resolved;
                    _attachmentAutomationSettingsSnapshotUtc =
                        DateTime.UtcNow;
                }
                return resolved;
            }

            private AttachmentAutomationSettings ReadLocalAttachmentAutomationSettings()
            {
                _owner.EnsureSettingsLoaded();

                var settings = _owner._currentSettings ?? new AddinSettings();
                int thresholdMb = OutlookAttachmentAutomationGuardService.NormalizeThresholdMb(settings.SharingAttachmentsOfferAboveMb);
                bool alwaysConnector = settings.SharingAttachmentsAlwaysConnector;
                bool offerAboveEnabled = settings.SharingAttachmentsOfferAboveEnabled && !alwaysConnector;
                return new AttachmentAutomationSettings
                {
                    AlwaysConnector = alwaysConnector,
                    OfferAboveEnabled = offerAboveEnabled,
                    ThresholdMb = thresholdMb,
                    ThresholdBytes = (long)thresholdMb * 1024L * 1024L
                };
            }

            private static AttachmentAutomationSettings ApplyAttachmentAutomationPolicy(
                AttachmentAutomationSettings local,
                BackendPolicyStatus policyStatus)
            {
                bool alwaysConnector = local != null && local.AlwaysConnector;
                bool offerAboveEnabled = local != null && local.OfferAboveEnabled;
                int thresholdMb = local != null
                    ? local.ThresholdMb
                    : AddinSettings.DefaultSharingAttachmentsOfferAboveMb;
                if (policyStatus != null && policyStatus.IsDomainActive("share"))
                {
                    bool policyBool;
                    int policyInt;

                    if (policyStatus.IsLocked("share", "attachments_always_via_ncconnector")
                        && policyStatus.TryGetPolicyBool("share", "attachments_always_via_ncconnector", out policyBool))
                    {
                        alwaysConnector = policyBool;
                    }
                    if (policyStatus.IsLocked("share", "attachments_min_size_mb"))
                    {
                        if (policyStatus.TryGetPolicyInt("share", "attachments_min_size_mb", out policyInt))
                        {
                            thresholdMb = OutlookAttachmentAutomationGuardService.NormalizeThresholdMb(policyInt);
                            offerAboveEnabled = true;
                        }
                        else if (policyStatus.HasPolicyKey("share", "attachments_min_size_mb"))
                        {
                            offerAboveEnabled = false;
                        }
                    }
                }
                return new AttachmentAutomationSettings
                {
                    AlwaysConnector = alwaysConnector,
                    OfferAboveEnabled = offerAboveEnabled,
                    ThresholdMb = thresholdMb,
                    ThresholdBytes = (long)thresholdMb * 1024L * 1024L
                };
            }

            private bool TryValidateAttachmentPolicyBeforeSend(ref bool cancel)
            {
                if (_disposed || cancel)
                {
                    return !cancel;
                }

                int attachmentCount = CountPolicyRelevantAttachments();
                if (attachmentCount <= 0)
                {
                    return true;
                }

                AttachmentAutomationSettings settings =
                    ReadAttachmentAutomationSettings();
                if (!settings.AlwaysConnector)
                {
                    return true;
                }

                OutlookAttachmentAutomationGuardService.GuardState guardState;
                if (_owner.TryGetAttachmentAutomationGuardState(
                    "send_gate",
                    _composeKey,
                    out guardState))
                {
                    cancel = true;
                    ShowForcedAttachmentProcessingError();
                    return false;
                }

                cancel = true;
                ShowForcedAttachmentProcessingError();
                LogFileLink(
                    "Compose send blocked by required attachment routing (composeKey="
                    + _composeKey
                    + ", remainingAttachments="
                    + attachmentCount.ToString(CultureInfo.InvariantCulture)
                    + ").");
                return false;
            }

            private static void ShowForcedAttachmentProcessingError()
            {
                MessageBox.Show(
                    Strings.FileLinkWizardAttachmentModeReasonAlways
                    + "\r\n\r\n"
                    + Strings.FileLinkWizardUploadFailed,
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }

        }
    }
}
