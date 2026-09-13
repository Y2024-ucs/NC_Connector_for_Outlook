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
            private sealed class BeforeAddShareEntry
            {
                internal AttachmentBatchEntry Candidate { get; set; }

                internal string LocalPath { get; set; }

                internal int ThresholdMb { get; set; }

                internal bool CleanupLocalPathAfterFlow { get; set; }

                internal string Trigger { get; set; }

                internal int BaselineAttachmentCount { get; set; }
            }

            private async Task StartComposeAttachmentShareFlowAsync(string trigger, long totalBytes, int thresholdMb, AttachmentBatchInfo lastAdded)
            {
                OutlookAttachmentAutomationGuardService.GuardState guardState;
                if (_owner.TryGetAttachmentAutomationGuardState("start_flow", _composeKey, out guardState))
                {
                    return;
                }
                var selections = new List<FileLinkSelection>();
                var removeIndices = new List<int>();
                var tempFiles = new List<string>();

                CollectAttachmentSelectionsForShare(selections, removeIndices, tempFiles);
                if (selections.Count == 0)
                {
                    CleanupTemporaryFiles(tempFiles);
                    LogFileLink("Compose attachment flow skipped (composeKey=" + _composeKey + ", reason=no_collectible_files).");
                    return;
                }

                var launchOptions = new FileLinkWizardLaunchOptions
                {
                    AttachmentMode = true,
                    AttachmentTrigger = string.IsNullOrWhiteSpace(trigger) ? "always" : trigger,
                    AttachmentTotalBytes = Math.Max(0, totalBytes),
                    AttachmentThresholdMb = Math.Max(1, thresholdMb),
                    AttachmentLastName = lastAdded != null ? (lastAdded.Name ?? string.Empty) : string.Empty,
                    AttachmentLastSizeBytes = lastAdded != null ? Math.Max(0, lastAdded.SizeBytes) : 0,
                    OnInitialQueueAdopted = () =>
                        RemoveAttachmentsByIndices(
                            removeIndices,
                            "share_flow")
                };
                for (int i = 0; i < selections.Count; i++)
                {
                    launchOptions.InitialSelections.Add(new FileLinkSelection(selections[i].SelectionType, selections[i].LocalPath));
                }
                try
                {
                    bool wizardAccepted = await _owner.RunFileLinkWizardForMailAsync(_mail, launchOptions);
                    LogFileLink(
                        "Compose attachment flow completed (composeKey="
                        + _composeKey
                        + ", trigger="
                        + launchOptions.AttachmentTrigger
                        + ", queued="
                        + selections.Count.ToString(CultureInfo.InvariantCulture)
                        + ", wizardAccepted="
                        + wizardAccepted.ToString(CultureInfo.InvariantCulture)
                        + ").");
                }
                finally
                {
                    CleanupTemporaryFiles(tempFiles);
                }
            }

            private void StartBeforeAddAttachmentShareFlow(
                string trigger,
                AttachmentBatchEntry candidate,
                string localPath,
                int thresholdMb,
                bool cleanupLocalPathAfterFlow)
            {
                int baselineAttachmentCount = ReadComposeAttachmentCount("before_add_single_baseline");
                RunAttachmentFlowTask(
                    StartBeforeAddAttachmentShareFlowAsync(trigger, candidate, localPath, thresholdMb, cleanupLocalPathAfterFlow, baselineAttachmentCount),
                    "Compose before-attachment-add share flow failed");
            }

            private async Task StartBeforeAddAttachmentShareFlowAsync(
                string trigger,
                AttachmentBatchEntry candidate,
                string localPath,
                int thresholdMb,
                bool cleanupLocalPathAfterFlow,
                int baselineAttachmentCount)
            {
                if (string.IsNullOrWhiteSpace(localPath) || !File.Exists(localPath))
                {
                    LogFileLink("Compose before-attachment-add share flow skipped (composeKey=" + _composeKey + ", reason=missing_local_path).");
                    return;
                }
                if (_beforeAddShareFlowRunning)
                {
                    QueueBeforeAddAttachmentShareFlow(trigger, candidate, localPath, thresholdMb, cleanupLocalPathAfterFlow);
                    return;
                }
                var launchOptions = new FileLinkWizardLaunchOptions
                {
                    AttachmentMode = true,
                    AttachmentTrigger = string.IsNullOrWhiteSpace(trigger) ? "threshold_preadd" : trigger,
                    AttachmentTotalBytes = Math.Max(0, candidate != null ? candidate.SizeBytes : 0),
                    AttachmentThresholdMb = Math.Max(1, thresholdMb),
                    AttachmentLastName = candidate != null ? (candidate.Name ?? string.Empty) : string.Empty,
                    AttachmentLastSizeBytes = Math.Max(0, candidate != null ? candidate.SizeBytes : 0)
                };
                launchOptions.InitialSelections.Add(new FileLinkSelection(FileLinkSelectionType.File, localPath));

                _beforeAddShareFlowRunning = true;
                _attachmentSuppressed = true;
                _pendingAddedBatch.Clear();
                try
                {
                    bool wizardAccepted = await _owner.RunFileLinkWizardForMailAsync(_mail, launchOptions);
                    LogFileLink(
                        "Compose before-attachment-add share flow completed (composeKey="
                        + _composeKey
                        + ", trigger="
                        + launchOptions.AttachmentTrigger
                        + ", wizardAccepted="
                        + wizardAccepted.ToString(CultureInfo.InvariantCulture)
                        + ", attachment="
                        + (candidate != null ? (candidate.Name ?? string.Empty) : string.Empty)
                        + ").");
                }
                finally
                {
                    RemoveSuppressedBeforeAddAttachmentByName(
                        candidate != null ? candidate.Name : string.Empty,
                        candidate != null ? candidate.SizeBytes : 0,
                        baselineAttachmentCount,
                        "before_add_single");
                    if (cleanupLocalPathAfterFlow)
                    {
                        CleanupTemporaryFiles(new List<string> { localPath });
                    }
                    EndAttachmentSuppression("before_add_single");
                    _beforeAddShareFlowRunning = false;
                    RestartBeforeAddShareTimerIfNeeded();
                }
            }

            private void QueueBeforeAddAttachmentShareFlow(
                string trigger,
                AttachmentBatchEntry candidate,
                string localPath,
                int thresholdMb,
                bool cleanupLocalPathAfterFlow)
            {
                if (string.IsNullOrWhiteSpace(localPath) || !File.Exists(localPath))
                {
                    LogFileLink("Compose before-attachment-add share queue skipped (composeKey=" + _composeKey + ", reason=missing_local_path).");
                    return;
                }

                _pendingBeforeAddShareEntries.Add(new BeforeAddShareEntry
                {
                    Candidate = new AttachmentBatchEntry
                    {
                        Name = candidate != null ? (candidate.Name ?? string.Empty) : string.Empty,
                        SizeBytes = Math.Max(0, candidate != null ? candidate.SizeBytes : 0)
                    },
                    LocalPath = localPath,
                    ThresholdMb = Math.Max(1, thresholdMb),
                    CleanupLocalPathAfterFlow = cleanupLocalPathAfterFlow,
                    Trigger = string.IsNullOrWhiteSpace(trigger) ? "always_preadd" : trigger,
                    BaselineAttachmentCount = ReadComposeAttachmentCount("before_add_queue_baseline")
                });

                LogFileLink(
                    "Compose before-attachment-add candidate queued (composeKey="
                    + _composeKey
                    + ", queued="
                    + _pendingBeforeAddShareEntries.Count.ToString(CultureInfo.InvariantCulture)
                    + ", attachment="
                    + (candidate != null ? (candidate.Name ?? string.Empty) : string.Empty)
                    + ").");

                if (_beforeAddShareFlowRunning)
                {
                    return;
                }
                try
                {
                    _beforeAddShareTimer.Stop();
                    _beforeAddShareTimer.Start();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to schedule queued before-attachment-add share flow (composeKey=" + _composeKey + ").",
                        ex);
                }
            }

            private async Task RunQueuedBeforeAddAttachmentShareFlowAsync()
            {
                if (_disposed || _beforeAddShareFlowRunning || _pendingBeforeAddShareEntries.Count == 0)
                {
                    return;
                }

                _beforeAddShareFlowRunning = true;
                var batch = new List<BeforeAddShareEntry>(_pendingBeforeAddShareEntries);
                _pendingBeforeAddShareEntries.Clear();

                var launchOptions = new FileLinkWizardLaunchOptions
                {
                    AttachmentMode = true,
                    AttachmentTrigger = "always_preadd_batch"
                };
                var temporaryFiles = new List<string>();
                long totalBytes = 0;
                int thresholdMb = 1;
                string lastName = string.Empty;
                long lastSize = 0;

                try
                {
                    for (int i = 0; i < batch.Count; i++)
                    {
                        BeforeAddShareEntry entry = batch[i];
                        if (entry == null || string.IsNullOrWhiteSpace(entry.LocalPath) || !File.Exists(entry.LocalPath))
                        {
                            continue;
                        }

                        launchOptions.InitialSelections.Add(new FileLinkSelection(FileLinkSelectionType.File, entry.LocalPath));
                        totalBytes += Math.Max(0, entry.Candidate != null ? entry.Candidate.SizeBytes : 0);
                        thresholdMb = Math.Max(thresholdMb, Math.Max(1, entry.ThresholdMb));
                        lastName = entry.Candidate != null ? (entry.Candidate.Name ?? string.Empty) : string.Empty;
                        lastSize = Math.Max(0, entry.Candidate != null ? entry.Candidate.SizeBytes : 0);
                        if (entry.CleanupLocalPathAfterFlow)
                        {
                            temporaryFiles.Add(entry.LocalPath);
                        }
                    }
                    if (launchOptions.InitialSelections.Count == 0)
                    {
                        LogFileLink("Compose before-attachment-add queued share flow skipped (composeKey=" + _composeKey + ", reason=no_collectible_files).");
                        return;
                    }

                    launchOptions.AttachmentTotalBytes = Math.Max(0, totalBytes);
                    launchOptions.AttachmentThresholdMb = thresholdMb;
                    launchOptions.AttachmentLastName = lastName;
                    launchOptions.AttachmentLastSizeBytes = lastSize;

                    _attachmentSuppressed = true;
                    _pendingAddedBatch.Clear();
                    bool wizardAccepted = await _owner.RunFileLinkWizardForMailAsync(_mail, launchOptions);
                    LogFileLink(
                        "Compose before-attachment-add queued share flow completed (composeKey="
                        + _composeKey
                        + ", queued="
                        + batch.Count.ToString(CultureInfo.InvariantCulture)
                        + ", selected="
                        + launchOptions.InitialSelections.Count.ToString(CultureInfo.InvariantCulture)
                        + ", wizardAccepted="
                        + wizardAccepted.ToString(CultureInfo.InvariantCulture)
                        + ").");
                }
                finally
                {
                    for (int i = 0; i < batch.Count; i++)
                    {
                        BeforeAddShareEntry entry = batch[i];
                        RemoveSuppressedBeforeAddAttachmentByName(
                            entry != null && entry.Candidate != null ? entry.Candidate.Name : string.Empty,
                            entry != null && entry.Candidate != null ? entry.Candidate.SizeBytes : 0,
                            entry != null ? entry.BaselineAttachmentCount : 0,
                            "before_add_batch");
                    }
                    CleanupTemporaryFiles(temporaryFiles);
                    EndAttachmentSuppression("before_add_batch");
                    _beforeAddShareFlowRunning = false;
                    RestartBeforeAddShareTimerIfNeeded();
                }
            }

            private void RunAttachmentFlowTask(Task task, string failureMessage)
            {
                if (task == null)
                {
                    return;
                }
                task.ContinueWith(
                    failedTask => DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        (failureMessage ?? "Compose attachment flow failed")
                        + " (composeKey="
                        + _composeKey
                        + ").",
                        failedTask.Exception),
                    TaskContinuationOptions.OnlyOnFaulted);
            }

            private void RestartBeforeAddShareTimerIfNeeded()
            {
                if (_pendingBeforeAddShareEntries.Count == 0 || _disposed)
                {
                    return;
                }
                try
                {
                    _beforeAddShareTimer.Stop();
                    _beforeAddShareTimer.Start();
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(
                        LogCategories.FileLink,
                        "Failed to restart queued before-attachment-add share flow (composeKey=" + _composeKey + ").",
                        ex);
                }
            }

            private void EndAttachmentSuppression(string trigger)
            {
                _attachmentSuppressed = false;
                ResumeDeferredEmailSignatureApplication("attachment_" + (trigger ?? string.Empty));
            }

        }
    }
}
