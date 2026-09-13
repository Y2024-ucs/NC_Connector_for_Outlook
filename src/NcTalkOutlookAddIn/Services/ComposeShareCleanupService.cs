// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Threading;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    internal sealed class ComposeShareCleanupService
    {
        internal bool TryDeleteComposeShareFolder(
            ComposeShareCleanupRecord entry,
            string reason)
        {
            if (entry == null)
            {
                return true;
            }
            string relativeFolder = entry.RelativeFolder;
            if (string.IsNullOrWhiteSpace(relativeFolder))
            {
                return true;
            }
            if (entry.Origin == null || !entry.Origin.IsComplete())
            {
                NextcloudTalkAddIn.LogFileLinkMessage(
                    "Compose share cleanup skipped (origin incomplete): relativeFolder="
                    + relativeFolder
                    + ", reason="
                    + (reason ?? string.Empty));
                return false;
            }
            TalkServiceConfiguration configuration =
                entry.Origin.ToConfiguration();
            var service = new FileLinkService(configuration);
            try
            {
                service.DeleteShareFolder(relativeFolder, CancellationToken.None);
                NextcloudTalkAddIn.LogFileLinkMessage(
                    "Compose share cleanup delete success (relativeFolder="
                    + relativeFolder
                    + ", reason="
                    + (reason ?? string.Empty)
                    + ", shareId="
                    + (entry.ShareId ?? string.Empty)
                    + ", shareLabel="
                    + (entry.ShareLabel ?? string.Empty)
                    + ").");
                return true;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Compose share cleanup delete failure (relativeFolder="
                    + relativeFolder
                    + ", reason="
                    + (reason ?? string.Empty)
                    + ", shareId="
                    + (entry.ShareId ?? string.Empty)
                    + ", shareLabel="
                    + (entry.ShareLabel ?? string.Empty)
                    + ").",
                    ex);
                return false;
            }
        }

    }
}
