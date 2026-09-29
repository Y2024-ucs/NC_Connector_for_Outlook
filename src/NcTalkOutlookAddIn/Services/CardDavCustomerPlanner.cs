// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace NcTalkOutlookAddIn.Services
{
    /// <summary>
    /// State of one contact in the Outlook customer folder, read on the Outlook UI thread.
    /// </summary>
    internal sealed class CardDavCustomerLocalEntry
    {
        internal string EntryId { get; set; }
        internal string Uid { get; set; }
        internal string Href { get; set; }
        internal string StoredEtag { get; set; }
        internal string StoredHash { get; set; }
        internal string CurrentHash { get; set; }
        internal CardDavContactRecord Fields { get; set; }

        internal bool LocallyChanged
        {
            get { return !string.Equals(CurrentHash ?? string.Empty, StoredHash ?? string.Empty, StringComparison.Ordinal); }
        }
    }

    internal enum CardDavCustomerActionKind
    {
        UploadNew,
        UploadUpdate,
        DownloadNew,
        DownloadUpdate,
        DeleteLocal,
        DeleteRemote
    }

    internal sealed class CardDavCustomerAction
    {
        internal CardDavCustomerActionKind Kind { get; set; }
        internal CardDavCustomerLocalEntry Local { get; set; }
        internal string Href { get; set; }
        internal string Uid { get; set; }
        internal string RemoteEtag { get; set; }

        // Outcome of the network phase, applied to Outlook afterwards.
        internal bool Succeeded { get; set; }
        internal CardDavContactRecord Downloaded { get; set; }
    }

    internal sealed class CardDavCustomerPlan
    {
        internal CardDavCustomerPlan()
        {
            Actions = new List<CardDavCustomerAction>();
            BlockedDeletionHrefs = new List<string>();
        }

        internal List<CardDavCustomerAction> Actions { get; private set; }
        internal List<string> BlockedDeletionHrefs { get; private set; }
        internal bool DeletionsBlocked { get { return BlockedDeletionHrefs.Count > 0; } }
    }

    /// <summary>
    /// Decides the two-way customer sync actions. Outlook wins when both sides changed.
    /// Deletions need evidence: the href must have been synced before (known) and be absent now.
    /// </summary>
    internal static class CardDavCustomerPlanner
    {
        internal const int MassDeletionMinimum = 5;

        internal static CardDavCustomerPlan Plan(
            IEnumerable<CardDavCustomerLocalEntry> localEntries,
            IDictionary<string, string> remoteVersions,
            ICollection<string> knownHrefs,
            string addressBookHref)
        {
            var plan = new CardDavCustomerPlan();
            var remote = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (KeyValuePair<string, string> pair in remoteVersions ?? new Dictionary<string, string>())
            {
                string href = NormalizeHref(pair.Key);
                if (href.Length > 0) remote[href] = pair.Value ?? string.Empty;
            }
            var known = new HashSet<string>(
                (knownHrefs ?? new string[0]).Select(NormalizeHref).Where(h => h.Length > 0),
                StringComparer.Ordinal);
            var claimed = new HashSet<string>(StringComparer.Ordinal);

            foreach (CardDavCustomerLocalEntry local in localEntries ?? Enumerable.Empty<CardDavCustomerLocalEntry>())
            {
                if (local == null) continue;
                string href = NormalizeHref(local.Href);

                // No marker, a foreign book or a copied contact sharing an href: upload as a new contact.
                if (href.Length == 0 || !BelongsTo(addressBookHref, href) || claimed.Contains(href))
                {
                    string uid = NewUid();
                    plan.Actions.Add(new CardDavCustomerAction
                    {
                        Kind = CardDavCustomerActionKind.UploadNew,
                        Local = local,
                        Uid = uid,
                        Href = BuildHref(addressBookHref, uid)
                    });
                    continue;
                }
                claimed.Add(href);

                string remoteEtag;
                if (remote.TryGetValue(href, out remoteEtag))
                {
                    if (local.LocallyChanged)
                    {
                        plan.Actions.Add(new CardDavCustomerAction
                        {
                            Kind = CardDavCustomerActionKind.UploadUpdate,
                            Local = local,
                            Href = href,
                            Uid = string.IsNullOrWhiteSpace(local.Uid) ? NewUid() : local.Uid.Trim(),
                            RemoteEtag = remoteEtag
                        });
                    }
                    else if (!EtagEquals(remoteEtag, local.StoredEtag))
                    {
                        plan.Actions.Add(new CardDavCustomerAction
                        {
                            Kind = CardDavCustomerActionKind.DownloadUpdate,
                            Local = local,
                            Href = href,
                            RemoteEtag = remoteEtag
                        });
                    }
                    continue;
                }

                if (known.Contains(href) && !local.LocallyChanged)
                {
                    plan.Actions.Add(new CardDavCustomerAction
                    {
                        Kind = CardDavCustomerActionKind.DeleteLocal,
                        Local = local,
                        Href = href
                    });
                }
                else
                {
                    // Edited in Outlook after a cloud deletion, or never confirmed: Outlook wins.
                    plan.Actions.Add(new CardDavCustomerAction
                    {
                        Kind = CardDavCustomerActionKind.UploadNew,
                        Local = local,
                        Href = href,
                        Uid = string.IsNullOrWhiteSpace(local.Uid) ? NewUid() : local.Uid.Trim()
                    });
                }
            }

            foreach (KeyValuePair<string, string> pair in remote.OrderBy(p => p.Key, StringComparer.Ordinal))
            {
                if (claimed.Contains(pair.Key)) continue;
                plan.Actions.Add(new CardDavCustomerAction
                {
                    Kind = known.Contains(pair.Key)
                        ? CardDavCustomerActionKind.DeleteRemote
                        : CardDavCustomerActionKind.DownloadNew,
                    Href = pair.Key,
                    RemoteEtag = pair.Value
                });
            }

            List<CardDavCustomerAction> deletions = plan.Actions
                .Where(a => a.Kind == CardDavCustomerActionKind.DeleteLocal
                    || a.Kind == CardDavCustomerActionKind.DeleteRemote)
                .ToList();
            if (deletions.Count >= MassDeletionMinimum && deletions.Count * 2 > known.Count)
            {
                foreach (CardDavCustomerAction deletion in deletions)
                {
                    plan.Actions.Remove(deletion);
                    plan.BlockedDeletionHrefs.Add(deletion.Href);
                }
            }
            return plan;
        }

        internal static bool BelongsTo(string addressBookHref, string href)
        {
            Uri book;
            Uri contact;
            if (!Uri.TryCreate((addressBookHref ?? string.Empty).TrimEnd('/') + "/", UriKind.Absolute, out book)
                || !Uri.TryCreate(href ?? string.Empty, UriKind.Absolute, out contact))
            {
                return false;
            }
            return string.IsNullOrEmpty(contact.Query)
                && string.IsNullOrEmpty(contact.Fragment)
                && !contact.AbsolutePath.EndsWith("/", StringComparison.Ordinal)
                && string.Equals(new Uri(contact, ".").AbsoluteUri, book.AbsoluteUri, StringComparison.Ordinal);
        }

        internal static string BuildHref(string addressBookHref, string uid)
        {
            string safe = new string((uid ?? string.Empty)
                .Where(c => (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || c == '-')
                .ToArray());
            if (safe.Length == 0) safe = Guid.NewGuid().ToString("D");
            return new Uri(new Uri((addressBookHref ?? string.Empty).TrimEnd('/') + "/"), safe + ".vcf").AbsoluteUri;
        }

        internal static string NormalizeHref(string href)
        {
            Uri uri;
            return Uri.TryCreate((href ?? string.Empty).Trim(), UriKind.Absolute, out uri)
                ? uri.AbsoluteUri
                : string.Empty;
        }

        internal static bool EtagEquals(string left, string right)
        {
            return string.Equals(NormalizeEtag(left), NormalizeEtag(right), StringComparison.Ordinal)
                && NormalizeEtag(left).Length > 0;
        }

        private static string NormalizeEtag(string value)
        {
            string result = (value ?? string.Empty).Trim();
            if (result.StartsWith("W/", StringComparison.Ordinal)) result = result.Substring(2);
            return result.Trim('"');
        }

        private static string NewUid()
        {
            return Guid.NewGuid().ToString("D");
        }
    }

    /// <summary>
    /// Builds vCards for customer contacts. Updates merge into the existing server vCard, so
    /// properties Outlook does not map (photo, birthday, URLs, ...) survive an upload.
    /// </summary>
    internal static class CardDavCustomerVCard
    {
        private static readonly HashSet<string> ManagedProperties = new HashSet<string>(
            new[] { "VERSION", "UID", "FN", "N", "ORG", "TITLE", "EMAIL", "TEL", "NOTE", "REV", "PRODID" },
            StringComparer.OrdinalIgnoreCase);

        internal static string ComputeHash(CardDavContactRecord record)
        {
            var builder = new StringBuilder();
            foreach (string value in FieldValues(record))
            {
                builder.Append(Clean(value).Replace("\r\n", "\n")).Append('\u001f');
            }
            using (SHA256 sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(builder.ToString()));
                return string.Concat(hash.Select(b => b.ToString("x2", CultureInfo.InvariantCulture)));
            }
        }

        internal static string Build(CardDavContactRecord record, string uid, string existingRaw, DateTime utcNow)
        {
            record = record ?? new CardDavContactRecord();
            var output = new List<string>();
            output.Add("BEGIN:VCARD");
            output.Add("VERSION:3.0");
            output.Add("PRODID:-//IBP//NC Connector for Outlook//DE");
            output.Add("UID:" + EscapeText(uid));
            string fullName = Clean(record.FullName);
            if (fullName.Length == 0)
            {
                fullName = (Clean(record.FirstName) + " " + Clean(record.LastName)).Trim();
            }
            if (fullName.Length == 0) fullName = Clean(record.Company);
            if (fullName.Length == 0) fullName = Clean(record.Email1);
            output.Add("FN:" + EscapeText(fullName));
            output.Add("N:" + EscapeText(record.LastName) + ";" + EscapeText(record.FirstName) + ";;;");
            AddIfPresent(output, "ORG", record.Company);
            AddIfPresent(output, "TITLE", record.JobTitle);
            AddIfPresent(output, "EMAIL;TYPE=INTERNET,WORK", record.Email1);
            AddIfPresent(output, "EMAIL;TYPE=INTERNET", record.Email2);
            AddIfPresent(output, "TEL;TYPE=WORK,VOICE", record.BusinessPhone);
            AddIfPresent(output, "TEL;TYPE=WORK,VOICE", record.BusinessPhone2);
            AddIfPresent(output, "TEL;TYPE=CELL", record.MobilePhone);
            AddIfPresent(output, "TEL;TYPE=HOME,VOICE", record.HomePhone);
            AddIfPresent(output, "TEL;TYPE=HOME,VOICE", record.HomePhone2);
            AddIfPresent(output, "TEL;TYPE=WORK,FAX", record.BusinessFax);
            AddIfPresent(output, "TEL;TYPE=HOME,FAX", record.HomeFax);
            AddIfPresent(output, "TEL;TYPE=OTHER", record.OtherPhone);
            AddIfPresent(output, "TEL;TYPE=PAGER", record.PagerPhone);
            AddIfPresent(output, "TEL;TYPE=MAIN", record.CompanyMainPhone);
            AddIfPresent(output, "TEL;TYPE=CAR", record.CarPhone);
            string[] address =
            {
                record.BusinessAddressStreet, record.BusinessAddressCity, record.BusinessAddressState,
                record.BusinessAddressPostalCode, record.BusinessAddressCountry
            };
            if (address.Any(part => Clean(part).Length > 0))
            {
                output.Add("ADR;TYPE=WORK:;;" + string.Join(";", address.Select(EscapeText)));
            }
            AddIfPresent(output, "NOTE", record.Notes);

            output.AddRange(PreservedLines(existingRaw));
            output.Add("REV:" + utcNow.ToUniversalTime().ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture));
            output.Add("END:VCARD");

            var builder = new StringBuilder();
            foreach (string line in output)
            {
                AppendFolded(builder, line);
            }
            return builder.ToString();
        }

        /// <summary>
        /// Lines of the server vCard that Outlook does not own. Business/untyped addresses and
        /// labels of dropped grouped properties are replaced; everything else is kept verbatim.
        /// </summary>
        private static IEnumerable<string> PreservedLines(string existingRaw)
        {
            if (string.IsNullOrWhiteSpace(existingRaw)) yield break;

            List<string> lines = Unfold(existingRaw);
            var droppedGroups = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string line in lines)
            {
                string group;
                string name = PropertyName(line, out group);
                if (group.Length > 0 && IsReplaced(name, line)) droppedGroups.Add(group);
            }

            int depth = 0;
            foreach (string line in lines)
            {
                if (line.Trim().Length == 0) continue;
                string group;
                string name = PropertyName(line, out group);
                string value = line.IndexOf(':') >= 0 ? line.Substring(line.IndexOf(':') + 1).Trim() : string.Empty;
                if (string.Equals(name, "BEGIN", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(value, "VCARD", StringComparison.OrdinalIgnoreCase))
                {
                    depth++;
                    continue;
                }
                if (string.Equals(name, "END", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(value, "VCARD", StringComparison.OrdinalIgnoreCase))
                {
                    depth--;
                    continue;
                }
                if (depth != 1 || IsReplaced(name, line)) continue;
                if (group.Length > 0 && droppedGroups.Contains(group)) continue;
                yield return line;
            }
        }

        private static bool IsReplaced(string name, string line)
        {
            if (ManagedProperties.Contains(name)) return true;
            if (!string.Equals(name, "ADR", StringComparison.OrdinalIgnoreCase)) return false;
            string parameters = line.IndexOf(':') >= 0 ? line.Substring(0, line.IndexOf(':')).ToUpperInvariant() : string.Empty;
            // Outlook only maps the business address; home and other addresses stay on the server.
            return !(parameters.Contains("HOME") || parameters.Contains("OTHER"));
        }

        private static string PropertyName(string line, out string group)
        {
            group = string.Empty;
            int colon = line.IndexOf(':');
            string left = colon >= 0 ? line.Substring(0, colon) : line;
            int semicolon = left.IndexOf(';');
            string token = (semicolon >= 0 ? left.Substring(0, semicolon) : left).Trim();
            int dot = token.LastIndexOf('.');
            if (dot > 0)
            {
                group = token.Substring(0, dot);
                token = token.Substring(dot + 1);
            }
            return token.ToUpperInvariant();
        }

        private static List<string> Unfold(string raw)
        {
            var result = new List<string>();
            foreach (string line in raw.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n'))
            {
                if ((line.StartsWith(" ", StringComparison.Ordinal) || line.StartsWith("\t", StringComparison.Ordinal))
                    && result.Count > 0)
                {
                    result[result.Count - 1] += line.Substring(1);
                }
                else
                {
                    result.Add(line);
                }
            }
            return result;
        }

        private static void AddIfPresent(List<string> output, string property, string value)
        {
            if (Clean(value).Length > 0) output.Add(property + ":" + EscapeText(value));
        }

        private static void AppendFolded(StringBuilder builder, string line)
        {
            const int limit = 74;
            int index = 0;
            bool first = true;
            while (index < line.Length)
            {
                int length = Math.Min(first ? limit : limit - 1, line.Length - index);
                if (index + length < line.Length && char.IsHighSurrogate(line[index + length - 1])) length--;
                if (!first) builder.Append(' ');
                builder.Append(line, index, length).Append("\r\n");
                index += length;
                first = false;
            }
            if (line.Length == 0) builder.Append("\r\n");
        }

        private static string EscapeText(string value)
        {
            return Clean(value)
                .Replace("\\", "\\\\")
                .Replace(";", "\\;")
                .Replace(",", "\\,")
                .Replace("\r\n", "\\n")
                .Replace("\r", "\\n")
                .Replace("\n", "\\n");
        }

        private static string Clean(string value)
        {
            return (value ?? string.Empty).Trim();
        }

        private static IEnumerable<string> FieldValues(CardDavContactRecord r)
        {
            r = r ?? new CardDavContactRecord();
            return new[]
            {
                r.FullName, r.FirstName, r.LastName, r.Company, r.JobTitle, r.Email1, r.Email2,
                r.BusinessPhone, r.BusinessPhone2, r.MobilePhone, r.HomePhone, r.HomePhone2,
                r.BusinessFax, r.HomeFax, r.OtherPhone, r.PagerPhone, r.CompanyMainPhone, r.CarPhone,
                r.BusinessAddressStreet, r.BusinessAddressCity, r.BusinessAddressState,
                r.BusinessAddressPostalCode, r.BusinessAddressCountry, r.Notes
            };
        }
    }
}
