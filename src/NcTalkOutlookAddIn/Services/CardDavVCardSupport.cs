// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Text;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.Services
{
    internal static class CardDavVCardSupport
    {
        private const long MaximumPhotoBytes = 5L * 1024L * 1024L;
        private const string GeneratedSystemAddressBookMarker = "z-server-generated--system";

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

        internal static void CachePhoto(
            string href,
            string rawVCard,
            TalkServiceConfiguration configuration)
        {
            if (string.IsNullOrWhiteSpace(href) || string.IsNullOrWhiteSpace(rawVCard))
            {
                return;
            }

            string uid = ExtractPropertyValue(rawVCard, "UID");
            string email = ExtractPropertyValue(rawVCard, "EMAIL");
            string photoValue;
            byte[] photo = TryExtractEmbeddedPhoto(rawVCard, out photoValue);
            string source = string.Empty;

            if (IsSupportedImage(photo))
            {
                source = "embedded";
            }
            else
            {
                photo = null;
            }

            if (photo == null
                && !string.IsNullOrWhiteSpace(photoValue)
                && configuration != null)
            {
                photo = TryDownloadPhoto(photoValue, configuration);
                if (IsSupportedImage(photo))
                {
                    source = "photo-uri";
                }
                else
                {
                    photo = null;
                }
            }

            if (photo == null
                && IsSystemDirectoryHref(href)
                && configuration != null)
            {
                string avatarIdentity;
                photo = TryDownloadAvatar(uid, email, configuration, out avatarIdentity);
                if (IsSupportedImage(photo))
                {
                    source = "nextcloud-avatar:" + avatarIdentity;
                }
                else
                {
                    photo = null;
                }
            }

            if (photo != null && photo.Length > 0)
            {
                Photos[href] = photo;
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "CardDAV contact photo cached (source="
                    + source
                    + ", uid="
                    + SafeLogValue(uid)
                    + ", bytes="
                    + photo.Length
                    + ").");
            }
            else if (IsSystemDirectoryHref(href))
            {
                DiagnosticsLogger.Log(
                    LogCategories.Core,
                    "CardDAV contact photo unavailable (uid="
                    + SafeLogValue(uid)
                    + ", email="
                    + SafeLogValue(email)
                    + ", hasPhotoField="
                    + (!string.IsNullOrWhiteSpace(photoValue))
                    + ").");
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
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "Failed to apply CardDAV contact photo in Outlook.",
                    ex);
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

        private static byte[] TryExtractEmbeddedPhoto(string rawVCard, out string photoValue)
        {
            photoValue = string.Empty;
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
            photoValue = value;
            if (value.Length == 0 || LooksLikePhotoUri(value))
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

        private static byte[] TryDownloadPhoto(
            string photoValue,
            TalkServiceConfiguration configuration)
        {
            try
            {
                string url = ResolvePhotoUrl(photoValue, configuration);
                return DownloadImage(url, configuration);
            }
            catch
            {
                return null;
            }
        }

        private static byte[] TryDownloadAvatar(
            string uid,
            string email,
            TalkServiceConfiguration configuration,
            out string matchedIdentity)
        {
            matchedIdentity = string.Empty;
            string baseUrl = configuration != null
                ? configuration.GetNormalizedBaseUrl()
                : string.Empty;
            if (string.IsNullOrWhiteSpace(baseUrl))
            {
                return null;
            }

            var identities = new List<string>();
            AddIdentity(identities, uid);
            if (!string.IsNullOrWhiteSpace(email))
            {
                string trimmedEmail = email.Trim();
                int at = trimmedEmail.IndexOf('@');
                if (at > 0)
                {
                    AddIdentity(identities, trimmedEmail.Substring(0, at));
                }
                AddIdentity(identities, trimmedEmail);
            }

            foreach (string identity in identities)
            {
                string escaped = Uri.EscapeDataString(identity);
                string[] urls =
                {
                    baseUrl.TrimEnd('/') + "/index.php/avatar/" + escaped + "/256",
                    baseUrl.TrimEnd('/') + "/avatar/" + escaped + "/256"
                };

                foreach (string url in urls)
                {
                    byte[] image = DownloadImage(url, configuration);
                    if (IsSupportedImage(image))
                    {
                        matchedIdentity = identity;
                        return image;
                    }
                }
            }

            return null;
        }

        private static byte[] DownloadImage(
            string url,
            TalkServiceConfiguration configuration)
        {
            if (string.IsNullOrWhiteSpace(url) || configuration == null)
            {
                return null;
            }

            try
            {
                NcHttpResponse response = new NcHttpClient(configuration).Send(new NcHttpRequestOptions
                {
                    Method = "GET",
                    Url = url,
                    Accept = "image/*, */*",
                    IncludeOcsApiHeader = false,
                    ParseJson = false,
                    ReadResponseAsBytes = true,
                    MaximumResponseBytes = MaximumPhotoBytes,
                    TimeoutMs = 30000
                });

                if (!response.HasHttpResponse
                    || (int)response.StatusCode < 200
                    || (int)response.StatusCode >= 300
                    || response.ResponseBytes == null
                    || response.ResponseBytes.Length == 0)
                {
                    return null;
                }

                return response.ResponseBytes;
            }
            catch
            {
                return null;
            }
        }

        private static string ResolvePhotoUrl(
            string value,
            TalkServiceConfiguration configuration)
        {
            string trimmed = (value ?? string.Empty).Trim();
            if (!LooksLikePhotoUri(trimmed))
            {
                return string.Empty;
            }

            Uri absolute;
            if (Uri.TryCreate(trimmed, UriKind.Absolute, out absolute))
            {
                if (string.Equals(absolute.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(absolute.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
                {
                    return absolute.AbsoluteUri;
                }
                return string.Empty;
            }

            string baseUrl = configuration != null
                ? configuration.GetNormalizedBaseUrl()
                : string.Empty;
            Uri baseUri;
            if (!Uri.TryCreate(baseUrl.TrimEnd('/') + "/", UriKind.Absolute, out baseUri))
            {
                return string.Empty;
            }

            Uri resolved;
            return Uri.TryCreate(baseUri, trimmed.TrimStart('/'), out resolved)
                ? resolved.AbsoluteUri
                : string.Empty;
        }

        private static string ExtractPropertyValue(string rawVCard, string propertyName)
        {
            if (string.IsNullOrWhiteSpace(rawVCard) || string.IsNullOrWhiteSpace(propertyName))
            {
                return string.Empty;
            }

            string normalized = rawVCard.Replace("\r\n", "\n").Replace("\r", "\n");
            string[] lines = normalized.Split('\n');
            foreach (string sourceLine in lines)
            {
                string line = sourceLine ?? string.Empty;
                int colon = line.IndexOf(':');
                if (colon <= 0)
                {
                    continue;
                }

                string left = line.Substring(0, colon);
                string name = left.Split(';')[0];
                int dot = name.LastIndexOf('.');
                if (dot >= 0 && dot + 1 < name.Length)
                {
                    name = name.Substring(dot + 1);
                }
                if (!string.Equals(name, propertyName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string value = line.Substring(colon + 1).Trim();
                if (string.Equals(propertyName, "EMAIL", StringComparison.OrdinalIgnoreCase)
                    && value.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase))
                {
                    value = value.Substring(7);
                }
                return value;
            }
            return string.Empty;
        }

        private static bool IsSystemDirectoryHref(string href)
        {
            return !string.IsNullOrWhiteSpace(href)
                && href.IndexOf(
                    GeneratedSystemAddressBookMarker,
                    StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static void AddIdentity(ICollection<string> identities, string value)
        {
            string candidate = (value ?? string.Empty).Trim();
            if (candidate.Length == 0)
            {
                return;
            }
            foreach (string existing in identities)
            {
                if (string.Equals(existing, candidate, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }
            }
            identities.Add(candidate);
        }

        private static bool LooksLikePhotoUri(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            return value.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
                || value.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
                || value.StartsWith("/", StringComparison.Ordinal)
                || value.StartsWith("./", StringComparison.Ordinal)
                || value.StartsWith("../", StringComparison.Ordinal);
        }

        private static bool IsSupportedImage(byte[] bytes)
        {
            if (bytes == null || bytes.Length < 2)
            {
                return false;
            }

            if (bytes[0] == 0xFF && bytes[1] == 0xD8)
            {
                return true;
            }
            if (bytes.Length >= 8
                && bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47)
            {
                return true;
            }
            if (bytes.Length >= 6
                && bytes[0] == 0x47 && bytes[1] == 0x49 && bytes[2] == 0x46)
            {
                return true;
            }
            return bytes[0] == 0x42 && bytes[1] == 0x4D;
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

        private static string SafeLogValue(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? "<empty>" : value.Trim();
        }
    }
}
