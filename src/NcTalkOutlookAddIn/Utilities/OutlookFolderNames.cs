// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;

namespace NcTalkOutlookAddIn.Utilities
{
    /// <summary>
    /// Compares Outlook folder names the way users see them. In IMAP profiles Outlook appends a
    /// "this computer only" suffix to local folders, so "Sifa" and "Sifa (Nur dieser Computer)" are one folder.
    /// </summary>
    internal static class OutlookFolderNames
    {
        private static readonly string[] LocalOnlySuffixes =
        {
            " (Nur dieser Computer)",
            " (This computer only)"
        };

        internal static string Normalize(string folderName)
        {
            string name = (folderName ?? string.Empty).Trim();
            foreach (string suffix in LocalOnlySuffixes)
            {
                if (name.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                {
                    return name.Substring(0, name.Length - suffix.Length).Trim();
                }
            }
            return name;
        }

        internal static bool Matches(string folderName, string wantedName)
        {
            return string.Equals(Normalize(folderName), Normalize(wantedName), StringComparison.OrdinalIgnoreCase);
        }
    }
}
