// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.


using System;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
    // Queue destination and Nextcloud storage summary.
    internal sealed partial class FileLinkWizardForm
    {
        private void RefreshQueueTargetPath()
        {
            FileLinkShareTarget target = FileLinkPath.ResolveShareTarget(
                _request.BasePath,
                _shareNameTextBox.Text,
                _shareDate,
                Strings.FileLinkWizardFallbackShareName);
            _basePathLabel.Text = "/" + target.RelativeFolderPath;
        }

        private void EnsureQueueStorageLoaded()
        {
            if (!IsHandleCreated
                || IsDisposed
                || Disposing
                || _queueStorageLoading
                || (_queueStorageTask != null
                    && !_queueStorageTask.IsCompleted)
                || _queueStorageLoaded)
            {
                return;
            }
            if (_queueStorageCancellation != null)
            {
                _queueStorageCancellation.Cancel();
                _queueStorageCancellation.Dispose();
            }
            _queueStorageCancellation = new CancellationTokenSource();
            CancellationToken token = _queueStorageCancellation.Token;
            int requestId = ++_queueStorageRequestId;
            _queueStorageLoading = true;
            RefreshQueueSummary();
            _queueStorageTask = RefreshQueueStorageAsync(
                requestId,
                token);
        }

        private async Task RefreshQueueStorageAsync(
            int requestId,
            CancellationToken cancellationToken)
        {
            NextcloudStorageListing listing = null;
            Exception failure = null;
            try
            {
                listing = await Task.Run(
                    () => _service.ListNextcloudDirectory(
                        string.Empty,
                        cancellationToken),
                    cancellationToken);
            }
            catch (OperationCanceledException)
            {
                return;
            }
            catch (Exception ex)
            {
                failure = ex;
            }

            if (requestId != _queueStorageRequestId
                || IsDisposed
                || Disposing)
            {
                return;
            }
            _queueStorageLoading = false;
            _queueStorageLoaded = true;
            _queueStorageUsedBytes = listing != null
                ? listing.UsedBytes
                : null;
            _queueStorageAvailableBytes = listing != null
                ? listing.AvailableBytes
                : null;
            if (failure != null)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Queue storage information could not be loaded.",
                    failure);
            }
            RefreshQueueSummary();
            UpdateUploadButtonState();
        }

        private void CancelQueueStorageRefresh()
        {
            _queueStorageRequestId++;
            if (_queueStorageCancellation == null)
            {
                return;
            }
            _queueStorageCancellation.Cancel();
            _queueStorageCancellation.Dispose();
            _queueStorageCancellation = null;
        }

        private void RefreshQueueSummary()
        {
            int entryCount = 0;
            long totalBytes = 0;
            foreach (FileLinkQueueNode root in _queueSnapshots.Values)
            {
                AccumulateQueueSummary(
                    root,
                    ref entryCount,
                    ref totalBytes);
            }
            int sourceCount = _items
                .Select(selection => selection.Source)
                .Distinct()
                .Count();
            _queueSummaryLabel.Text = string.Join(
                " · ",
                new[]
                {
                    string.Format(
                        CultureInfo.CurrentCulture,
                        Strings.FileLinkQueueEntriesSummary,
                        entryCount),
                    string.Format(
                        CultureInfo.CurrentCulture,
                        Strings.FileLinkQueueSourcesSummary,
                        sourceCount),
                    string.Format(
                        CultureInfo.CurrentCulture,
                        Strings.FileLinkQueueTotalSummary,
                        SizeFormatting.FormatBytes(totalBytes))
                });

            _queueStorageLabel.ForeColor = _themePalette.MutedText;
            if (_queueStorageLoading)
            {
                _queueStorageLabel.Text =
                    Strings.FileLinkQueueStorageLoading;
                return;
            }
            if (!_queueStorageLoaded || !_queueStorageUsedBytes.HasValue)
            {
                _queueStorageLabel.Text =
                    Strings.FileLinkQueueStorageUnknown;
                return;
            }
            if (!_queueStorageAvailableBytes.HasValue)
            {
                _queueStorageLabel.Text = string.Format(
                    CultureInfo.CurrentCulture,
                    Strings.FileLinkQueueStorageUnlimited,
                    SizeFormatting.FormatBytes(
                        _queueStorageUsedBytes.Value));
                return;
            }
            if (totalBytes > _queueStorageAvailableBytes.Value)
            {
                _queueStorageLabel.ForeColor = _themePalette.ErrorText;
                _queueStorageLabel.Text = string.Format(
                    CultureInfo.CurrentCulture,
                    Strings.FileLinkQueueStorageInsufficient,
                    SizeFormatting.FormatBytes(totalBytes),
                    SizeFormatting.FormatBytes(
                        _queueStorageAvailableBytes.Value));
                return;
            }

            long quota = _queueStorageUsedBytes.Value
                         > long.MaxValue
                           - _queueStorageAvailableBytes.Value
                ? long.MaxValue
                : _queueStorageUsedBytes.Value
                  + _queueStorageAvailableBytes.Value;
            _queueStorageLabel.Text = string.Format(
                CultureInfo.CurrentCulture,
                Strings.FileLinkQueueStorageAvailable,
                SizeFormatting.FormatBytes(
                    _queueStorageAvailableBytes.Value),
                SizeFormatting.FormatBytes(quota));
        }

        private static void AccumulateQueueSummary(
            FileLinkQueueNode node,
            ref int entryCount,
            ref long totalBytes)
        {
            if (node == null)
            {
                return;
            }
            if (entryCount < int.MaxValue)
            {
                entryCount++;
            }
            if (!node.IsDirectory && node.Length.HasValue)
            {
                totalBytes = totalBytes > long.MaxValue - node.Length.Value
                    ? long.MaxValue
                    : totalBytes + node.Length.Value;
            }
            foreach (FileLinkQueueNode child in node.Children)
            {
                AccumulateQueueSummary(
                    child,
                    ref entryCount,
                    ref totalBytes);
            }
        }

        private bool HasInsufficientQueueStorage()
        {
            if (!_queueStorageLoaded
                || !_queueStorageAvailableBytes.HasValue)
            {
                return false;
            }
            long totalBytes = 0;
            int ignoredCount = 0;
            foreach (FileLinkQueueNode root in _queueSnapshots.Values)
            {
                AccumulateQueueSummary(
                    root,
                    ref ignoredCount,
                    ref totalBytes);
            }
            return totalBytes > _queueStorageAvailableBytes.Value;
        }
    }
}
