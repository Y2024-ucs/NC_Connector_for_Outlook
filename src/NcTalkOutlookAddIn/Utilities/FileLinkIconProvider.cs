// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Windows.Forms;

namespace NcTalkOutlookAddIn.Utilities
{
    internal static class FileLinkIconProvider
    {
        internal const string LocalSourceKey = "source-local";
        internal const string NextcloudSourceKey = "source-nextcloud";
        internal const string FolderKey = "folder";
        internal const string GenericFileKey = "file";

        private enum FileIconKind
        {
            Generic,
            Memo,
            Web,
            Palette,
            Script,
            Clipboard,
            Chart,
            RedBook,
            BlueBook,
            GreenBook,
            OrangeBook,
            Archive,
            Picture,
            Audio,
            Video
        }

        private static readonly IDictionary<string, FileIconKind>
            ExtensionIcons =
            new Dictionary<string, FileIconKind>(
                StringComparer.OrdinalIgnoreCase)
            {
                { ".txt", FileIconKind.Generic },
                { ".md", FileIconKind.Memo },
                { ".html", FileIconKind.Web },
                { ".htm", FileIconKind.Web },
                { ".css", FileIconKind.Palette },
                { ".js", FileIconKind.Script },
                { ".mjs", FileIconKind.Script },
                { ".ts", FileIconKind.Script },
                { ".json", FileIconKind.Clipboard },
                { ".xml", FileIconKind.Clipboard },
                { ".csv", FileIconKind.Chart },
                { ".pdf", FileIconKind.RedBook },
                { ".doc", FileIconKind.BlueBook },
                { ".docx", FileIconKind.BlueBook },
                { ".odt", FileIconKind.BlueBook },
                { ".xls", FileIconKind.GreenBook },
                { ".xlsx", FileIconKind.GreenBook },
                { ".ods", FileIconKind.GreenBook },
                { ".ppt", FileIconKind.OrangeBook },
                { ".pptx", FileIconKind.OrangeBook },
                { ".odp", FileIconKind.OrangeBook },
                { ".zip", FileIconKind.Archive },
                { ".7z", FileIconKind.Archive },
                { ".rar", FileIconKind.Archive },
                { ".tar", FileIconKind.Archive },
                { ".gz", FileIconKind.Archive },
                { ".png", FileIconKind.Picture },
                { ".jpg", FileIconKind.Picture },
                { ".jpeg", FileIconKind.Picture },
                { ".gif", FileIconKind.Picture },
                { ".svg", FileIconKind.Picture },
                { ".webp", FileIconKind.Picture },
                { ".mp3", FileIconKind.Audio },
                { ".wav", FileIconKind.Audio },
                { ".ogg", FileIconKind.Audio },
                { ".mp4", FileIconKind.Video },
                { ".webm", FileIconKind.Video },
                { ".mkv", FileIconKind.Video },
                { ".mov", FileIconKind.Video },
                { ".avi", FileIconKind.Video }
            };

        internal static ImageList CreateImageList(
            int width,
            int height)
        {
            int safeWidth = Math.Max(16, width);
            int safeHeight = Math.Max(20, height);
            var images = new ImageList
            {
                ColorDepth = ColorDepth.Depth32Bit,
                ImageSize = new Size(safeWidth, safeHeight),
                TransparentColor = Color.Transparent
            };
            images.Images.Add(
                LocalSourceKey,
                DrawLocalSource(safeWidth, safeHeight));
            images.Images.Add(
                NextcloudSourceKey,
                DrawNextcloudSource(safeWidth, safeHeight));
            images.Images.Add(
                FolderKey,
                DrawFolder(safeWidth, safeHeight));
            images.Images.Add(
                GenericFileKey,
                DrawFile(
                    safeWidth,
                    safeHeight,
                    FileIconKind.Generic));
            foreach (KeyValuePair<string, FileIconKind> pair
                in ExtensionIcons)
            {
                images.Images.Add(
                    BuildExtensionKey(pair.Key),
                    DrawFile(safeWidth, safeHeight, pair.Value));
            }
            return images;
        }

        internal static string GetFileKey(string fileName)
        {
            string extension = Path.GetExtension(fileName ?? string.Empty);
            return !string.IsNullOrWhiteSpace(extension)
                   && ExtensionIcons.ContainsKey(extension)
                ? BuildExtensionKey(extension)
                : GenericFileKey;
        }

        private static string BuildExtensionKey(string extension)
        {
            return "file-" + (extension ?? string.Empty)
                .TrimStart('.')
                .ToLowerInvariant();
        }

        private static Bitmap CreateCanvas(int width, int height)
        {
            return new Bitmap(width, height);
        }

        private static Bitmap DrawFolder(int width, int height)
        {
            Bitmap bitmap = CreateCanvas(width, height);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle bounds = CenteredBounds(width, height, 17, 14);
                using (var back = new SolidBrush(
                    Color.FromArgb(226, 163, 45)))
                using (var front = new SolidBrush(
                    Color.FromArgb(255, 190, 57)))
                {
                    graphics.FillRectangle(
                        back,
                        bounds.Left + 1,
                        bounds.Top,
                        7,
                        4);
                    graphics.FillRectangle(
                        front,
                        bounds.Left,
                        bounds.Top + 3,
                        bounds.Width,
                        bounds.Height - 3);
                }
            }
            return bitmap;
        }

        private static Bitmap DrawFile(
            int width,
            int height,
            FileIconKind kind)
        {
            switch (kind)
            {
                case FileIconKind.Memo:
                    return DrawMemo(width, height);
                case FileIconKind.Web:
                    return DrawGlobe(width, height);
                case FileIconKind.Palette:
                    return DrawPalette(width, height);
                case FileIconKind.Script:
                    return DrawScript(width, height);
                case FileIconKind.Clipboard:
                    return DrawClipboard(width, height);
                case FileIconKind.Chart:
                    return DrawChart(width, height);
                case FileIconKind.RedBook:
                    return DrawBook(
                        width,
                        height,
                        Color.FromArgb(224, 47, 84));
                case FileIconKind.BlueBook:
                    return DrawBook(
                        width,
                        height,
                        Color.FromArgb(47, 170, 222));
                case FileIconKind.GreenBook:
                    return DrawBook(
                        width,
                        height,
                        Color.FromArgb(113, 202, 87));
                case FileIconKind.OrangeBook:
                    return DrawBook(
                        width,
                        height,
                        Color.FromArgb(235, 93, 42));
                case FileIconKind.Archive:
                    return DrawArchive(width, height);
                case FileIconKind.Picture:
                    return DrawPicture(width, height);
                case FileIconKind.Audio:
                    return DrawAudio(width, height);
                case FileIconKind.Video:
                    return DrawVideo(width, height);
                default:
                    return DrawDocument(width, height);
            }
        }

        private static Bitmap DrawDocument(int width, int height)
        {
            Bitmap bitmap = CreateCanvas(width, height);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle bounds = CenteredBounds(width, height, 13, 17);
                var outline = new[]
                {
                    new Point(bounds.Left, bounds.Top),
                    new Point(bounds.Right - 5, bounds.Top),
                    new Point(bounds.Right, bounds.Top + 5),
                    new Point(bounds.Right, bounds.Bottom),
                    new Point(bounds.Left, bounds.Bottom)
                };
                using (var fill = new SolidBrush(
                    Color.FromArgb(238, 235, 244)))
                using (var border = new Pen(
                    Color.FromArgb(190, 184, 198)))
                using (var line = new Pen(
                    Color.FromArgb(178, 169, 191)))
                {
                    graphics.FillPolygon(fill, outline);
                    graphics.DrawPolygon(border, outline);
                    graphics.DrawLine(
                        border,
                        bounds.Right - 5,
                        bounds.Top,
                        bounds.Right - 5,
                        bounds.Top + 5);
                    graphics.DrawLine(
                        border,
                        bounds.Right - 5,
                        bounds.Top + 5,
                        bounds.Right,
                        bounds.Top + 5);
                    for (int offset = 8; offset <= 13; offset += 3)
                    {
                        graphics.DrawLine(
                            line,
                            bounds.Left + 3,
                            bounds.Top + offset,
                            bounds.Right - 3,
                            bounds.Top + offset);
                    }
                }
            }
            return bitmap;
        }

        private static Bitmap DrawMemo(int width, int height)
        {
            Bitmap bitmap = DrawDocument(width, height);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            using (var pencil = new Pen(
                Color.FromArgb(237, 116, 46),
                2.6f))
            using (var tip = new Pen(
                Color.FromArgb(99, 76, 63),
                1.2f))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle bounds = CenteredBounds(width, height, 13, 17);
                graphics.DrawLine(
                    pencil,
                    bounds.Left + 4,
                    bounds.Bottom - 3,
                    bounds.Right - 1,
                    bounds.Top + 5);
                graphics.DrawLine(
                    tip,
                    bounds.Left + 3,
                    bounds.Bottom - 2,
                    bounds.Left + 5,
                    bounds.Bottom - 4);
            }
            return bitmap;
        }

        private static Bitmap DrawBook(
            int width,
            int height,
            Color color)
        {
            Bitmap bitmap = CreateCanvas(width, height);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle bounds = CenteredBounds(width, height, 13, 17);
                using (var cover = new SolidBrush(color))
                using (var spine = new SolidBrush(
                    ControlPaint.Dark(color, 0.22f)))
                using (var edge = new Pen(
                    ControlPaint.Dark(color, 0.12f)))
                using (var pages = new Pen(
                    Color.FromArgb(220, 245, 245, 245)))
                {
                    graphics.FillRectangle(cover, bounds);
                    graphics.FillRectangle(
                        spine,
                        bounds.Left,
                        bounds.Top,
                        2,
                        bounds.Height);
                    graphics.DrawRectangle(
                        edge,
                        bounds.Left,
                        bounds.Top,
                        bounds.Width - 1,
                        bounds.Height - 1);
                    graphics.DrawLine(
                        pages,
                        bounds.Left + 3,
                        bounds.Top + 2,
                        bounds.Right - 2,
                        bounds.Top + 2);
                    graphics.DrawLine(
                        pages,
                        bounds.Left + 3,
                        bounds.Bottom - 2,
                        bounds.Right - 2,
                        bounds.Bottom - 2);
                }
            }
            return bitmap;
        }

        private static Bitmap DrawGlobe(int width, int height)
        {
            Bitmap bitmap = CreateCanvas(width, height);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            using (var outline = new Pen(
                Color.FromArgb(43, 183, 232),
                1.4f))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle bounds = CenteredBounds(width, height, 17, 17);
                graphics.DrawEllipse(outline, bounds);
                graphics.DrawEllipse(
                    outline,
                    bounds.Left + 5,
                    bounds.Top,
                    bounds.Width - 10,
                    bounds.Height);
                graphics.DrawLine(
                    outline,
                    bounds.Left + 1,
                    bounds.Top + bounds.Height / 2,
                    bounds.Right - 1,
                    bounds.Top + bounds.Height / 2);
                graphics.DrawArc(
                    outline,
                    bounds.Left + 1,
                    bounds.Top + 3,
                    bounds.Width - 2,
                    bounds.Height - 6,
                    190,
                    160);
            }
            return bitmap;
        }

        private static Bitmap DrawPalette(int width, int height)
        {
            Bitmap bitmap = CreateCanvas(width, height);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle bounds = CenteredBounds(width, height, 18, 16);
                using (var body = new SolidBrush(
                    Color.FromArgb(242, 177, 132)))
                using (var hole = new SolidBrush(Color.Transparent))
                {
                    graphics.FillEllipse(body, bounds);
                    graphics.CompositingMode = CompositingMode.SourceCopy;
                    graphics.FillEllipse(
                        hole,
                        bounds.Right - 7,
                        bounds.Bottom - 7,
                        5,
                        5);
                    graphics.CompositingMode =
                        CompositingMode.SourceOver;
                }
                DrawPaletteDot(
                    graphics,
                    bounds.Left + 4,
                    bounds.Top + 4,
                    Color.FromArgb(223, 69, 83));
                DrawPaletteDot(
                    graphics,
                    bounds.Left + 9,
                    bounds.Top + 3,
                    Color.FromArgb(63, 152, 230));
                DrawPaletteDot(
                    graphics,
                    bounds.Left + 5,
                    bounds.Top + 9,
                    Color.FromArgb(96, 190, 85));
            }
            return bitmap;
        }

        private static void DrawPaletteDot(
            Graphics graphics,
            int left,
            int top,
            Color color)
        {
            using (var brush = new SolidBrush(color))
            {
                graphics.FillEllipse(brush, left, top, 3, 3);
            }
        }

        private static Bitmap DrawScript(int width, int height)
        {
            Bitmap bitmap = CreateCanvas(width, height);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle bounds = CenteredBounds(width, height, 14, 17);
                using (var paper = new SolidBrush(
                    Color.FromArgb(245, 193, 130)))
                using (var roll = new Pen(
                    Color.FromArgb(197, 117, 61),
                    2f))
                using (var line = new Pen(
                    Color.FromArgb(139, 94, 69)))
                {
                    graphics.FillRectangle(
                        paper,
                        bounds.Left + 1,
                        bounds.Top + 1,
                        bounds.Width - 2,
                        bounds.Height - 2);
                    graphics.DrawLine(
                        roll,
                        bounds.Left,
                        bounds.Top + 1,
                        bounds.Right,
                        bounds.Top + 1);
                    graphics.DrawLine(
                        roll,
                        bounds.Left,
                        bounds.Bottom - 1,
                        bounds.Right,
                        bounds.Bottom - 1);
                    graphics.DrawLine(
                        line,
                        bounds.Left + 4,
                        bounds.Top + 6,
                        bounds.Right - 3,
                        bounds.Top + 6);
                    graphics.DrawLine(
                        line,
                        bounds.Left + 4,
                        bounds.Top + 10,
                        bounds.Right - 5,
                        bounds.Top + 10);
                }
            }
            return bitmap;
        }

        private static Bitmap DrawClipboard(int width, int height)
        {
            Bitmap bitmap = CreateCanvas(width, height);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle bounds = CenteredBounds(width, height, 13, 17);
                using (var board = new SolidBrush(
                    Color.FromArgb(207, 151, 96)))
                using (var paper = new SolidBrush(
                    Color.FromArgb(242, 239, 245)))
                using (var clip = new SolidBrush(
                    Color.FromArgb(142, 135, 148)))
                using (var line = new Pen(
                    Color.FromArgb(174, 166, 181)))
                {
                    graphics.FillRectangle(board, bounds);
                    graphics.FillRectangle(
                        paper,
                        bounds.Left + 2,
                        bounds.Top + 3,
                        bounds.Width - 4,
                        bounds.Height - 5);
                    graphics.FillRectangle(
                        clip,
                        bounds.Left + 4,
                        bounds.Top,
                        bounds.Width - 8,
                        4);
                    graphics.DrawLine(
                        line,
                        bounds.Left + 4,
                        bounds.Top + 8,
                        bounds.Right - 3,
                        bounds.Top + 8);
                    graphics.DrawLine(
                        line,
                        bounds.Left + 4,
                        bounds.Top + 12,
                        bounds.Right - 4,
                        bounds.Top + 12);
                }
            }
            return bitmap;
        }

        private static Bitmap DrawChart(int width, int height)
        {
            Bitmap bitmap = CreateCanvas(width, height);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                Rectangle bounds = CenteredBounds(width, height, 17, 16);
                using (var frame = new SolidBrush(
                    Color.FromArgb(238, 235, 244)))
                using (var border = new Pen(
                    Color.FromArgb(173, 166, 181)))
                using (var blue = new SolidBrush(
                    Color.FromArgb(55, 157, 221)))
                using (var green = new SolidBrush(
                    Color.FromArgb(89, 190, 108)))
                using (var pink = new SolidBrush(
                    Color.FromArgb(218, 80, 137)))
                {
                    graphics.FillRectangle(frame, bounds);
                    graphics.DrawRectangle(
                        border,
                        bounds.Left,
                        bounds.Top,
                        bounds.Width - 1,
                        bounds.Height - 1);
                    int bottom = bounds.Bottom - 2;
                    graphics.FillRectangle(
                        blue,
                        bounds.Left + 3,
                        bottom - 6,
                        3,
                        6);
                    graphics.FillRectangle(
                        green,
                        bounds.Left + 7,
                        bottom - 10,
                        3,
                        10);
                    graphics.FillRectangle(
                        pink,
                        bounds.Left + 11,
                        bottom - 8,
                        3,
                        8);
                }
            }
            return bitmap;
        }

        private static Bitmap DrawArchive(int width, int height)
        {
            Bitmap bitmap = CreateCanvas(width, height);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                Rectangle bounds = CenteredBounds(width, height, 14, 17);
                using (var body = new SolidBrush(
                    Color.FromArgb(190, 146, 79)))
                using (var border = new Pen(
                    Color.FromArgb(116, 96, 83)))
                using (var zipper = new Pen(
                    Color.FromArgb(238, 232, 218)))
                {
                    graphics.FillRectangle(body, bounds);
                    graphics.DrawRectangle(
                        border,
                        bounds.Left,
                        bounds.Top,
                        bounds.Width - 1,
                        bounds.Height - 1);
                    int center = bounds.Left + bounds.Width / 2;
                    graphics.DrawLine(
                        zipper,
                        center,
                        bounds.Top + 1,
                        center,
                        bounds.Bottom - 2);
                    for (int top = bounds.Top + 2;
                        top < bounds.Bottom - 2;
                        top += 4)
                    {
                        graphics.DrawLine(
                            border,
                            center - 2,
                            top,
                            center + 2,
                            top);
                    }
                }
            }
            return bitmap;
        }

        private static Bitmap DrawPicture(int width, int height)
        {
            Bitmap bitmap = CreateCanvas(width, height);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle bounds = CenteredBounds(width, height, 18, 15);
                using (var frame = new SolidBrush(
                    Color.FromArgb(226, 137, 84)))
                using (var sky = new SolidBrush(
                    Color.FromArgb(93, 192, 218)))
                using (var hill = new SolidBrush(
                    Color.FromArgb(91, 175, 91)))
                using (var sun = new SolidBrush(
                    Color.FromArgb(248, 210, 68)))
                {
                    graphics.FillRectangle(frame, bounds);
                    var inside = new Rectangle(
                        bounds.Left + 2,
                        bounds.Top + 2,
                        bounds.Width - 4,
                        bounds.Height - 4);
                    graphics.FillRectangle(sky, inside);
                    graphics.FillEllipse(
                        sun,
                        inside.Right - 5,
                        inside.Top + 1,
                        3,
                        3);
                    graphics.FillPolygon(
                        hill,
                        new[]
                        {
                            new Point(inside.Left, inside.Bottom),
                            new Point(
                                inside.Left + inside.Width / 2,
                                inside.Top + 4),
                            new Point(inside.Right, inside.Bottom)
                        });
                }
            }
            return bitmap;
        }

        private static Bitmap DrawAudio(int width, int height)
        {
            Bitmap bitmap = CreateCanvas(width, height);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            using (var pen = new Pen(
                Color.FromArgb(105, 82, 162),
                2.8f))
            using (var fill = new SolidBrush(
                Color.FromArgb(105, 82, 162)))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle bounds = CenteredBounds(width, height, 16, 17);
                graphics.DrawLine(
                    pen,
                    bounds.Left + 7,
                    bounds.Top + 2,
                    bounds.Left + 7,
                    bounds.Bottom - 4);
                graphics.DrawLine(
                    pen,
                    bounds.Left + 7,
                    bounds.Top + 2,
                    bounds.Right - 2,
                    bounds.Top);
                graphics.DrawLine(
                    pen,
                    bounds.Right - 2,
                    bounds.Top,
                    bounds.Right - 2,
                    bounds.Bottom - 7);
                graphics.FillEllipse(
                    fill,
                    bounds.Left + 1,
                    bounds.Bottom - 6,
                    7,
                    5);
                graphics.FillEllipse(
                    fill,
                    bounds.Right - 8,
                    bounds.Bottom - 9,
                    7,
                    5);
            }
            return bitmap;
        }

        private static Bitmap DrawVideo(int width, int height)
        {
            Bitmap bitmap = CreateCanvas(width, height);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle bounds = CenteredBounds(width, height, 18, 16);
                using (var body = new SolidBrush(
                    Color.FromArgb(101, 79, 156)))
                using (var top = new SolidBrush(
                    Color.FromArgb(217, 207, 241)))
                using (var stripe = new Pen(
                    Color.FromArgb(101, 79, 156),
                    2f))
                {
                    graphics.FillRectangle(
                        body,
                        bounds.Left,
                        bounds.Top + 5,
                        bounds.Width,
                        bounds.Height - 5);
                    graphics.FillRectangle(
                        top,
                        bounds.Left,
                        bounds.Top,
                        bounds.Width,
                        5);
                    for (int left = bounds.Left - 2;
                        left < bounds.Right;
                        left += 6)
                    {
                        graphics.DrawLine(
                            stripe,
                            left,
                            bounds.Top + 4,
                            left + 4,
                            bounds.Top);
                    }
                }
            }
            return bitmap;
        }

        private static Bitmap DrawNextcloudSource(
            int width,
            int height)
        {
            Bitmap bitmap = CreateCanvas(width, height);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                Image appIcon = BrandingAssets.AppIconPng;
                if (appIcon != null)
                {
                    graphics.CompositingQuality =
                        CompositingQuality.HighQuality;
                    graphics.InterpolationMode =
                        InterpolationMode.HighQualityBicubic;
                    graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    int size = Math.Min(
                        Math.Min(width, height),
                        20);
                    Rectangle bounds = CenteredBounds(
                        width,
                        height,
                        size,
                        size);
                    graphics.DrawImage(appIcon, bounds);
                }
            }
            return bitmap;
        }

        private static Bitmap DrawLocalSource(
            int width,
            int height)
        {
            Bitmap bitmap = CreateCanvas(width, height);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle bounds = CenteredBounds(width, height, 17, 14);
                using (var pen = new Pen(
                    Color.FromArgb(0, 130, 201),
                    1.8f))
                {
                    graphics.DrawRectangle(
                        pen,
                        bounds.Left,
                        bounds.Top,
                        bounds.Width - 1,
                        bounds.Height - 4);
                    graphics.DrawLine(
                        pen,
                        bounds.Left + 5,
                        bounds.Bottom - 1,
                        bounds.Right - 5,
                        bounds.Bottom - 1);
                    graphics.DrawLine(
                        pen,
                        bounds.Left + 8,
                        bounds.Bottom - 4,
                        bounds.Left + 8,
                        bounds.Bottom - 1);
                }
            }
            return bitmap;
        }

        private static Rectangle CenteredBounds(
            int width,
            int height,
            int desiredWidth,
            int desiredHeight)
        {
            int actualWidth = Math.Min(width, desiredWidth);
            int actualHeight = Math.Min(height, desiredHeight);
            return new Rectangle(
                Math.Max(0, (width - actualWidth) / 2),
                Math.Max(0, (height - actualHeight) / 2),
                actualWidth,
                actualHeight);
        }
    }
}
