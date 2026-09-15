// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

namespace NcTalkOutlookAddIn.Models
{
    internal sealed class CalDavCalendar
    {
        internal string Href { get; set; }
        internal string DisplayName { get; set; }
        internal string Description { get; set; }
        internal string SyncToken { get; set; }
        internal string ETag { get; set; }
        internal bool ReadOnly { get; set; }

        public override string ToString()
        {
            return string.IsNullOrWhiteSpace(DisplayName) ? Href : DisplayName;
        }
    }
}
