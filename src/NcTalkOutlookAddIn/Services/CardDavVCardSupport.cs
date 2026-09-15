// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Concurrent;
using System.IO;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Services
{
    internal static class CardDavVCardSupport
    {
        private static readonly ConcurrentDictionary<string, byte[]> Photos =
            new ConcurrentDictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);

        internal static string Normalize(string raw)
        {
            if (string.IsNullOrEmpty(raw))
            {
                return string.Empty;
            }

            string normalized = raw.Replace("\r\n", "\n").Replace("\r", "\n");
            string[] lines = normalized.Split('\n');
            var builder = new StringBuilder();
            foreach (string sourceLine in lines)
            {
                string line = sourceLine ?? string.Empty;
                int colon = line.IndexOf(':');
                if (colon > 0)
                {
                    string left = line.Substring(0, colon);
                    string value = line.Substring(colon + 1);
                    int semicolon = left.IndexOf(';');
                    string propertyName = semicolon >= 0 ? left.Substring(0, semicolon) : left;
                    string parameters = semicolon >= 0 ? left.Substring(semicolon) : string.Empty;
                    int dot = propertyName.LastIndexOf('.');
                    if (dot >= 0 && dot + 1 < propertyName.Length)
                    {
                        propertyName = propertyName.Substring(dot + 1);
                    }

                    if (string.Equals(propertyName, "EMAIL", StringComparison.OrdinalIgnoreCase)
                        && value.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase))
                    {
                        value = value.Substring(7);
                    }

                    line = propertyName + parameters + ":" + value;
                }

                if (builder.Length > 0)
                {
                    builder.Append("\r\n");
                }
                builder.Append(line);
            }
            return builder.ToString();
        }

        internal static void CachePhoto(string href, string rawVCard)
        {
            if (string.IsNullOrWhiteSpace(href) || string.IsNullOrWhiteSpace(rawVCard))
            {
                return;
            }

            byte[] photo = TryExtractPhoto(rawVCard);
            if (photo != null && photo.Length > 0)
            {
                Photos[href] = photo;
            }
        }

        internal static void ApplyPhoto(Outlook.ContactItem target, string href)
        {
            if (target == null || string.IsNullOrWhiteSpace(href))
            {
                return;
            }

            byte[] photo;
            if (!Photos.TryGetValue(href, out photo) || photo == null || photo.Length == 0)
            {
                return;
            }

            string tempPath = null;
            try
            {
                tempPath = Path.Combine(
                    Path.GetTempPath(),
                    "nc4ol-contact-" + Guid.NewGuid().ToString("N") + DetectImageExtension(photo));
                File.WriteAllBytes(tempPath, photo);
                target.AddPicture(tempPath);
            }
            catch
            {
                // A broken or unsupported image must not abort the contact sync.
            }
            finally
            {
                if (!string.IsNullOrWhiteSpace(tempPath))
                {
                    try
                    {
                        if (File.Exists(tempPath))
                        {
                            File.Delete(tempPath);
                        }
                    }
                    catch
                    {
                    }
                }
            }
        }

        private static byte[] TryExtractPhoto(string rawVCard)
        {
            string normalized = rawVCard.Replace("\r\n", "\n").Replace("\r", "\n");
            string[] lines = normalized.Split('\n');
            var logical = new StringBuilder();
            bool collecting = false;

            foreach (string sourceLine in lines)
            {
                string line = sourceLine ?? string.Empty;
                if (!collecting)
                {
                    int colon = line.IndexOf(':');
                    if (colon <= 0)
                    {
                        continue;
                    }

                    string left = line.Substring(0, colon);
                    string propertyName = left.Split(';')[0];
                    int dot = propertyName.LastIndexOf('.');
                    if (dot >= 0 && dot + 1 < propertyName.Length)
                    {
                        propertyName = propertyName.Substring(dot + 1);
                    }
                    if (!string.Equals(propertyName, "PHOTO", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    collecting = true;
                    logical.Append(line.Substring(colon + 1).Trim());
                    continue;
                }

                if (line.StartsWith(" ") || line.StartsWith("\t"))
                {
                    logical.Append(line.Trim());
                    continue;
                }
                break;
            }

            string value = logical.ToString().Trim();
            if (value.Length == 0 || value.StartsWith("http", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }
            if (value.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
            {
                int comma = value.IndexOf(',');
                if (comma >= 0)
                {
                    value = value.Substring(comma + 1);
                }
            }

            try
            {
                return Convert.FromBase64String(value);
            }
            catch
            {
                return null;
            }
        }

        private static string DetectImageExtension(byte[] bytes)
        {
            if (bytes != null && bytes.Length >= 8
                && bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47)
            {
                return ".png";
            }
            if (bytes != null && bytes.Length >= 6
                && bytes[0] == 0x47 && bytes[1] == 0x49 && bytes[2] == 0x46)
            {
                return ".gif";
            }
            if (bytes != null && bytes.Length >= 2
                && bytes[0] == 0x42 && bytes[1] == 0x4D)
            {
                return ".bmp";
            }
            return ".jpg";
        }
    }
}
