// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.IO;

namespace NcTalkOutlookAddIn.Utilities
{
    internal static class WordHtmlInsertionFile
    {
        internal static void TryDeleteTemporaryHtmlFile(string tempHtmlPath, string failureMessage)
        {
            if (string.IsNullOrWhiteSpace(tempHtmlPath))
            {
                return;
            }

            try
            {
                File.Delete(tempHtmlPath);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, failureMessage, ex);
            }
        }

        internal static string EnsureHtmlDocumentForWordInsert(string html)
        {
            string value = html ?? string.Empty;
            if (value.IndexOf("<html", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return value;
            }

            return "<html><head><meta charset=\"utf-8\"><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"></head><body>"
                   + value
                   + "</body></html>";
        }

    }
}
