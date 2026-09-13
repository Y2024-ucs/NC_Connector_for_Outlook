// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;

namespace NcTalkOutlookAddIn.Utilities
{
    internal static class RecipientAddressList
    {
        internal static void AddUniqueRecipient(List<string> recipients, string address)
        {
            if (recipients == null)
            {
                return;
            }
            string normalized = NormalizeRecipientAddress(address);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return;
            }
            for (int i = 0; i < recipients.Count; i++)
            {
                if (string.Equals(recipients[i], normalized, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }
            }

            recipients.Add(normalized);
        }

        internal static List<string> ExtractRecipientAddresses(string csv)
        {
            var list = new List<string>();
            if (string.IsNullOrWhiteSpace(csv))
            {
                return list;
            }
            string[] parts = csv.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < parts.Length; i++)
            {
                AddUniqueRecipient(list, parts[i]);
            }
            return list;
        }

        internal static string BuildNormalizedRecipientCsv(string csv)
        {
            List<string> recipients = ExtractRecipientAddresses(csv);
            return recipients.Count == 0 ? string.Empty : string.Join("; ", recipients.ToArray());
        }

        internal static int CountRecipientsInCsv(string csv)
        {
            return ExtractRecipientAddresses(csv).Count;
        }

        internal static string NormalizeRecipientAddress(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                return string.Empty;
            }
            string value = raw.Trim();
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }
            int lt = value.LastIndexOf('<');
            int gt = value.LastIndexOf('>');
            if (lt >= 0 && gt > lt)
            {
                value = value.Substring(lt + 1, gt - lt - 1).Trim();
            }

            value = value.Trim().Trim('\'', '"');
            return value.Trim();
        }

    }
}
