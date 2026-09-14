// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Threading;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;

namespace NcTalkOutlookAddIn.UI
{
    internal sealed class NextcloudPickerPreview
    {
        private const long MaximumPreviewBytes = 5L * 1024L * 1024L;
        private const long MaximumDecodedPreviewPixels = 80000000L;
        private const int MaximumPreviewDimension = 2048;
        private const int RequestedGeneratedPreviewDimension = 1024;

        private static readonly ISet<string> PreviewImageExtensions =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                ".bmp",
                ".gif",
                ".ico",
                ".jpeg",
                ".jpg",
                ".png"
            };

        private readonly FileLinkService _service;
        private readonly NextcloudStorageEntry _entry;
        private readonly bool _supportsOriginalImagePreview;

        internal bool OriginalImageExceedsLimit { get; private set; }

        internal NextcloudPickerPreview(FileLinkService service, NextcloudStorageEntry entry)
        {
            _service = service;
            _entry = entry;
            _supportsOriginalImagePreview = PreviewImageExtensions.Contains(
                Path.GetExtension(entry.DisplayName) ?? string.Empty);
            OriginalImageExceedsLimit = _supportsOriginalImagePreview
                && entry.Length > MaximumPreviewBytes;
        }

        internal Bitmap Load(CancellationToken token)
        {
            byte[] bytes =
                _service.TryReadNextcloudGeneratedPreview(
                    _entry,
                    RequestedGeneratedPreviewDimension,
                    RequestedGeneratedPreviewDimension,
                    MaximumPreviewBytes,
                    token);
            if (bytes == null
                && _supportsOriginalImagePreview
                && !OriginalImageExceedsLimit)
            {
                bytes = _service.ReadNextcloudFilePreview(
                    _entry,
                    MaximumPreviewBytes,
                    token);
            }
            if (bytes == null)
            {
                return null;
            }
            token.ThrowIfCancellationRequested();
            Bitmap bitmap = DecodePreviewImage(bytes);
            token.ThrowIfCancellationRequested();
            return bitmap;
        }

        private static Bitmap DecodePreviewImage(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0)
            {
                throw new InvalidDataException(
                    "The preview response did not contain image data.");
            }

            using (var stream = new MemoryStream(bytes, false))
            using (Image source = Image.FromStream(
                stream,
                true,
                true))
            {
                long pixelCount = checked(
                    (long)source.Width * source.Height);
                if (source.Width <= 0
                    || source.Height <= 0
                    || pixelCount > MaximumDecodedPreviewPixels)
                {
                    throw new InvalidDataException(
                        "The image dimensions exceed the preview limit.");
                }

                double scale = Math.Min(
                    1d,
                    Math.Min(
                        (double)MaximumPreviewDimension / source.Width,
                        (double)MaximumPreviewDimension / source.Height));
                int width = Math.Max(
                    1,
                    (int)Math.Round(source.Width * scale));
                int height = Math.Max(
                    1,
                    (int)Math.Round(source.Height * scale));
                var bitmap = new Bitmap(
                    width,
                    height,
                    PixelFormat.Format32bppPArgb);
                try
                {
                    using (Graphics graphics = Graphics.FromImage(bitmap))
                    {
                        graphics.Clear(Color.Transparent);
                        graphics.CompositingQuality =
                            CompositingQuality.HighQuality;
                        graphics.InterpolationMode =
                            InterpolationMode.HighQualityBicubic;
                        graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                        graphics.SmoothingMode = SmoothingMode.HighQuality;
                        graphics.DrawImage(
                            source,
                            new Rectangle(0, 0, width, height));
                    }
                    return bitmap;
                }
                catch
                {
                    bitmap.Dispose();
                    throw;
                }
            }
        }

    }
}
