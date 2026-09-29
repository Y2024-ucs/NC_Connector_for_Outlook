// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

namespace NcTalkOutlookAddIn.Models
{
    internal sealed class CardDavAddressBook
    {
        internal string Href { get; set; }
        internal string DisplayName { get; set; }
        internal string Description { get; set; }
        internal string SyncToken { get; set; }
        internal string ETag { get; set; }

        internal CardDavAddressBook()
        {
            Href = string.Empty;
            DisplayName = string.Empty;
            Description = string.Empty;
            SyncToken = string.Empty;
            ETag = string.Empty;
        }
    }
}
