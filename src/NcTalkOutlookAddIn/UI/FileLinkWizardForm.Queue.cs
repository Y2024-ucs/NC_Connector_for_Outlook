// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.


using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
    // Queue tree interaction, rendering, and upload-state presentation.
    internal sealed partial class FileLinkWizardForm
    {
        private void RebuildQueueView()
        {
            if (_fileListView == null
                || _fileListView.IsDisposed
                || _fileListView.Disposing)
            {
                return;
            }

            foreach (SelectionUploadState state in _selectionStates.Values)
            {
                state.Item = null;
                if (state.ProgressBar != null)
                {
                    state.ProgressBar.Visible = false;
                }
            }

            _fileListView.BeginUpdate();
            try
            {
                _fileListView.Items.Clear();
                AddQueueSourceSection(FileLinkSelectionSource.Local);
                AddQueueSourceSection(FileLinkSelectionSource.Nextcloud);
            }
            finally
            {
                _fileListView.EndUpdate();
            }

            _queueEmptyLabel.Visible = _items.Count == 0;
            if (_queueEmptyLabel.Visible)
            {
                _queueEmptyLabel.BringToFront();
            }
            else
            {
                _fileListView.BringToFront();
            }
            RefreshQueueSummary();
            UpdateQueueColumnWidths();
            PositionProgressBars();
        }

        private void AddQueueSourceSection(
            FileLinkSelectionSource source)
        {
            List<FileLinkSelection> selections = _items
                .Where(selection => selection.Source == source)
                .ToList();
            if (selections.Count == 0)
            {
                return;
            }

            string label = source == FileLinkSelectionSource.Local
                ? Strings.FileLinkQueueSourceLocal
                : Strings.FileLinkQueueSourceNextcloud;
            var sourceRow = new FileLinkQueueRow(
                FileLinkQueueRowKind.Source,
                source,
                null,
                null,
                0,
                string.Empty,
                false,
                false);
            var sourceItem = new ListViewItem(label)
            {
                Tag = sourceRow,
                UseItemStyleForSubItems = false
            };
            sourceItem.SubItems.Add(string.Empty);
            sourceItem.SubItems.Add(string.Empty);
            sourceItem.SubItems.Add(string.Empty);
            _fileListView.Items.Add(sourceItem);

            foreach (FileLinkSelection selection in selections)
            {
                FileLinkQueueNode root;
                if (!_queueSnapshots.TryGetValue(selection, out root)
                    || root == null)
                {
                    continue;
                }
                AddQueueNode(selection, root, 0, true);
            }
        }

        private void AddQueueNode(
            FileLinkSelection selection,
            FileLinkQueueNode node,
            int depth,
            bool canRemove)
        {
            string key = BuildQueueNodeKey(selection, node);
            bool expanded = node.IsDirectory
                            && _expandedQueueFolders.Contains(key);
            var row = new FileLinkQueueRow(
                node.IsDirectory
                    ? FileLinkQueueRowKind.Folder
                    : FileLinkQueueRowKind.File,
                selection.Source,
                selection,
                node,
                depth,
                key,
                expanded,
                canRemove);
            string iconKey = node.IsDirectory
                ? FileLinkIconProvider.FolderKey
                : FileLinkIconProvider.GetFileKey(node.DisplayName);
            var item = new ListViewItem(node.DisplayName, iconKey)
            {
                Tag = row,
                ToolTipText = node.DisplayName,
                UseItemStyleForSubItems = false
            };
            item.SubItems.Add(
                node.IsDirectory || !node.Length.HasValue
                    ? string.Empty
                    : SizeFormatting.FormatBytes(node.Length.Value));
            item.SubItems.Add(string.Empty);
            item.SubItems.Add(string.Empty);
            _fileListView.Items.Add(item);

            SelectionUploadState state;
            if (_selectionStates.TryGetValue(selection, out state))
            {
                if (state.Item == null
                    || (!node.IsDirectory
                        && ResolveQueueRow(state.Item).Kind
                        == FileLinkQueueRowKind.Folder))
                {
                    state.Item = item;
                }
                ApplyQueueItemState(item, state);
            }

            if (!expanded)
            {
                return;
            }
            foreach (FileLinkQueueNode child in node.Children)
            {
                AddQueueNode(selection, child, depth + 1, false);
            }
        }

        private void ApplyQueueItemState(
            ListViewItem item,
            SelectionUploadState state)
        {
            FileLinkQueueRow row = ResolveQueueRow(item);
            if (row == null || row.Kind == FileLinkQueueRowKind.Source)
            {
                return;
            }
            string status = string.Empty;
            Color statusColor = _themePalette.MutedText;
            if (row.Kind == FileLinkQueueRowKind.File)
            {
                switch (state.Status)
                {
                    case FileLinkUploadStatus.Completed:
                        status = Strings.FileLinkWizardStatusSuccess;
                        statusColor = _themePalette.SuccessText;
                        break;
                    case FileLinkUploadStatus.Failed:
                        status = Strings.FileLinkWizardStatusError;
                        statusColor = _themePalette.ErrorText;
                        break;
                    default:
                        status = Strings.FileLinkQueueWaiting;
                        break;
                }
            }
            item.SubItems[2].Text = status;
            item.SubItems[2].ForeColor = statusColor;
        }

        private static string BuildQueueNodeKey(
            FileLinkSelection selection,
            FileLinkQueueNode node)
        {
            return selection.IdentityPath
                   + "|"
                   + (node != null ? node.SourcePath : string.Empty);
        }

        private void ToggleQueueFolder(FileLinkQueueRow row)
        {
            if (row == null || !row.HasChildren || IsWizardBusy)
            {
                return;
            }
            if (!_expandedQueueFolders.Remove(row.Key))
            {
                _expandedQueueFolders.Add(row.Key);
            }
            RebuildQueueView();
        }

        private void HandleQueueMouseClick(
            object sender,
            MouseEventArgs e)
        {
            if (e == null)
            {
                return;
            }
            ListViewHitTestInfo hit = _fileListView.HitTest(e.Location);
            if (hit.Item == null)
            {
                return;
            }
            FileLinkQueueRow row = ResolveQueueRow(hit.Item);
            if (row == null || row.Kind == FileLinkQueueRowKind.Source)
            {
                return;
            }
            int columnIndex = hit.SubItem == null
                ? -1
                : hit.Item.SubItems.IndexOf(hit.SubItem);
            if (columnIndex == 3 && row.CanRemove)
            {
                RemoveSelections(new[] { row.Selection });
                return;
            }
            if (columnIndex == 0
                && row.HasChildren
                && GetQueueToggleBounds(hit.Item, row).Contains(e.Location))
            {
                ToggleQueueFolder(row);
            }
        }

        private void HandleQueueMouseDoubleClick(
            object sender,
            MouseEventArgs e)
        {
            if (e == null)
            {
                return;
            }
            ListViewHitTestInfo hit = _fileListView.HitTest(e.Location);
            FileLinkQueueRow row = hit.Item != null
                ? ResolveQueueRow(hit.Item)
                : null;
            if (row != null && row.HasChildren)
            {
                ToggleQueueFolder(row);
            }
        }

        private void HandleQueueMouseMove(
            object sender,
            MouseEventArgs e)
        {
            if (e == null)
            {
                return;
            }
            ListViewHitTestInfo hit = _fileListView.HitTest(e.Location);
            FileLinkQueueRow row = hit.Item != null
                ? ResolveQueueRow(hit.Item)
                : null;
            string hint = string.Empty;
            if (row != null && row.Node != null)
            {
                int columnIndex = hit.SubItem == null
                    ? -1
                    : hit.Item.SubItems.IndexOf(hit.SubItem);
                if (columnIndex == 3 && row.CanRemove)
                {
                    hint = string.Format(
                        CultureInfo.CurrentCulture,
                        Strings.FileLinkQueueRemoveItem,
                        row.Node.DisplayName);
                }
                else if (columnIndex == 0 && row.HasChildren)
                {
                    hint = string.Format(
                        CultureInfo.CurrentCulture,
                        row.Expanded
                            ? Strings.FileLinkQueueCollapseFolder
                            : Strings.FileLinkQueueExpandFolder,
                        row.Node.DisplayName);
                }
                else
                {
                    hint = row.Node.DisplayName;
                }
            }
            if (!string.Equals(
                _toolTip.GetToolTip(_fileListView),
                hint,
                StringComparison.CurrentCulture))
            {
                _toolTip.SetToolTip(_fileListView, hint);
            }
        }

        private void HandleQueueKeyDown(object sender, KeyEventArgs e)
        {
            if (e == null)
            {
                return;
            }
            if (e.KeyCode == Keys.Delete)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                RemoveSelection();
                return;
            }
            if (_fileListView.SelectedItems.Count != 1)
            {
                return;
            }
            FileLinkQueueRow row = ResolveQueueRow(
                _fileListView.SelectedItems[0]);
            if (row == null || !row.HasChildren)
            {
                return;
            }
            if (e.KeyCode == Keys.Space
                || e.KeyCode == Keys.Enter
                || (e.KeyCode == Keys.Right && !row.Expanded)
                || (e.KeyCode == Keys.Left && row.Expanded))
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                ToggleQueueFolder(row);
            }
        }

        private Rectangle GetQueueToggleBounds(
            ListViewItem item,
            FileLinkQueueRow row)
        {
            Rectangle bounds = item.GetBounds(ItemBoundsPortion.Entire);
            int size = ScaleLogical(22);
            return new Rectangle(
                ScaleLogical(4) + (row.Depth * ScaleLogical(FileQueueIndentPixels)),
                bounds.Top + Math.Max(0, (bounds.Height - size) / 2),
                size,
                size);
        }

        private void UpdateQueueColumnWidths()
        {
            if (_fileListView == null
                || _fileListView.Columns.Count < 4
                || _fileListView.IsDisposed
                || _fileListView.Disposing)
            {
                return;
            }
            int clientWidth = Math.Max(0, _fileListView.ClientSize.Width);
            int sizeWidth = ScaleLogical(96);
            int statusWidth = ScaleLogical(120);
            int actionWidth = ScaleLogical(38);
            int nameWidth = Math.Max(
                ScaleLogical(150),
                clientWidth - sizeWidth - statusWidth - actionWidth - 6);
            _fileListView.Columns[0].Width = nameWidth;
            _fileListView.Columns[1].Width = sizeWidth;
            _fileListView.Columns[2].Width = statusWidth;
            _fileListView.Columns[3].Width = actionWidth;
            _fileListView.Invalidate();
        }

        private void HandleFileListViewDrawColumnHeader(
            object sender,
            DrawListViewColumnHeaderEventArgs e)
        {
            if (e != null)
            {
                e.DrawDefault = true;
            }
        }

        private void HandleFileListViewDrawItem(
            object sender,
            DrawListViewItemEventArgs e)
        {
            if (e != null && _fileListView.View != View.Details)
            {
                e.DrawDefault = true;
            }
        }

        private void HandleFileListViewDrawSubItem(
            object sender,
            DrawListViewSubItemEventArgs e)
        {
            if (e == null || e.Item == null || e.SubItem == null)
            {
                return;
            }
            FileLinkQueueRow row = ResolveQueueRow(e.Item);
            if (row == null)
            {
                e.DrawDefault = true;
                return;
            }

            bool selected = row.Kind != FileLinkQueueRowKind.Source
                            && e.Item.Selected
                            && (!_fileListView.HideSelection
                                || _fileListView.Focused);
            Color rowBackground = e.Item.BackColor.IsEmpty
                ? _themePalette.InputBackground
                : e.Item.BackColor;
            Color background = row.Kind == FileLinkQueueRowKind.Source
                ? _themePalette.ControlBackground
                : selected
                    ? _themePalette.SelectionBackground
                    : rowBackground;
            Color foreground = selected
                ? _themePalette.SelectionText
                : row.Kind == FileLinkQueueRowKind.Source
                    ? _themePalette.Text
                    : e.SubItem.ForeColor.IsEmpty
                        ? _themePalette.Text
                        : e.SubItem.ForeColor;
            using (var brush = new SolidBrush(background))
            {
                e.Graphics.FillRectangle(brush, e.Bounds);
            }

            if (row.Kind == FileLinkQueueRowKind.Source)
            {
                if (e.ColumnIndex == 0)
                {
                    DrawQueueSourceHeading(e, row, foreground);
                }
            }
            else if (e.ColumnIndex == 0)
            {
                DrawQueueNameCell(e, row, foreground);
            }
            else if (e.ColumnIndex == 1)
            {
                DrawQueueTextCell(
                    e,
                    foreground,
                    TextFormatFlags.Right);
            }
            else if (e.ColumnIndex == 2)
            {
                DrawQueueTextCell(
                    e,
                    foreground,
                    TextFormatFlags.Left);
            }
            else if (e.ColumnIndex == 3 && row.CanRemove)
            {
                DrawQueueRemoveGlyph(e.Graphics, e.Bounds, foreground);
            }

            using (var pen = new Pen(_themePalette.Border))
            {
                e.Graphics.DrawLine(
                    pen,
                    e.Bounds.Left,
                    e.Bounds.Bottom - 1,
                    e.Bounds.Right,
                    e.Bounds.Bottom - 1);
            }
            if (e.ColumnIndex == _fileListView.Columns.Count - 1
                && selected)
            {
                Rectangle focus = e.Item.Bounds;
                focus.Width = Math.Max(
                    0,
                    _fileListView.ClientSize.Width - focus.Left);
                ControlPaint.DrawFocusRectangle(
                    e.Graphics,
                    focus,
                    foreground,
                    background);
            }
        }

        private void DrawQueueSourceHeading(
            DrawListViewSubItemEventArgs e,
            FileLinkQueueRow row,
            Color foreground)
        {
            string imageKey = row.Source == FileLinkSelectionSource.Local
                ? FileLinkIconProvider.LocalSourceKey
                : FileLinkIconProvider.NextcloudSourceKey;
            Image image = _fileQueueImageList.Images[imageKey];
            int imageLeft = e.Bounds.Left + ScaleLogical(8);
            int imageTop = e.Bounds.Top
                           + Math.Max(0, (e.Bounds.Height - image.Height) / 2);
            e.Graphics.DrawImageUnscaled(image, imageLeft, imageTop);
            var textBounds = new Rectangle(
                imageLeft + image.Width + ScaleLogical(5),
                e.Bounds.Top,
                Math.Max(
                    0,
                    e.Bounds.Width - image.Width - ScaleLogical(18)),
                e.Bounds.Height);
            using (var font = new Font(
                _fileListView.Font,
                FontStyle.Bold))
            {
                TextRenderer.DrawText(
                    e.Graphics,
                    e.SubItem.Text,
                    font,
                    textBounds,
                    foreground,
                    TextFormatFlags.Left
                    | TextFormatFlags.VerticalCenter
                    | TextFormatFlags.EndEllipsis
                    | TextFormatFlags.NoPadding);
            }
        }

        private void DrawQueueNameCell(
            DrawListViewSubItemEventArgs e,
            FileLinkQueueRow row,
            Color foreground)
        {
            int indent = row.Depth * ScaleLogical(FileQueueIndentPixels);
            Rectangle toggleBounds = GetQueueToggleBounds(e.Item, row);
            if (row.HasChildren)
            {
                string arrow = row.Expanded ? "▾" : "▸";
                TextRenderer.DrawText(
                    e.Graphics,
                    arrow,
                    _fileListView.Font,
                    toggleBounds,
                    _themePalette.MutedText,
                    TextFormatFlags.HorizontalCenter
                    | TextFormatFlags.VerticalCenter
                    | TextFormatFlags.NoPadding);
            }

            if (row.Depth > 0)
            {
                using (var pen = new Pen(_themePalette.Border))
                {
                    int lineX = e.Bounds.Left
                                + indent
                                - ScaleLogical(7);
                    e.Graphics.DrawLine(
                        pen,
                        lineX,
                        e.Bounds.Top,
                        lineX,
                        e.Bounds.Bottom);
                }
            }

            string imageKey = row.Kind == FileLinkQueueRowKind.Folder
                ? FileLinkIconProvider.FolderKey
                : FileLinkIconProvider.GetFileKey(
                    row.Node != null
                        ? row.Node.DisplayName
                        : string.Empty);
            Image image = _fileQueueImageList.Images[imageKey];
            int imageLeft = e.Bounds.Left
                            + indent
                            + ScaleLogical(27);
            int imageTop = e.Bounds.Top
                           + Math.Max(0, (e.Bounds.Height - image.Height) / 2);
            e.Graphics.DrawImageUnscaled(image, imageLeft, imageTop);
            var textBounds = new Rectangle(
                imageLeft + image.Width + ScaleLogical(5),
                e.Bounds.Top,
                Math.Max(
                    0,
                    e.Bounds.Right
                    - imageLeft
                    - image.Width
                    - ScaleLogical(8)),
                e.Bounds.Height);
            TextRenderer.DrawText(
                e.Graphics,
                e.SubItem.Text,
                _fileListView.Font,
                textBounds,
                foreground,
                TextFormatFlags.Left
                | TextFormatFlags.VerticalCenter
                | TextFormatFlags.EndEllipsis
                | TextFormatFlags.NoPadding);
        }

        private void DrawQueueTextCell(
            DrawListViewSubItemEventArgs e,
            Color foreground,
            TextFormatFlags alignment)
        {
            int top = e.Bounds.Top;
            int height = e.Bounds.Height;
            SelectionUploadState state = e.ColumnIndex == 2
                ? ResolveSelectionState(e.Item)
                : null;
            if (state != null
                && state.Status == FileLinkUploadStatus.Uploading
                && ReferenceEquals(state.Item, e.Item))
            {
                int progressSpace = ScaleLogical(10);
                top += progressSpace;
                height = Math.Max(0, height - progressSpace);
            }
            var bounds = new Rectangle(
                e.Bounds.Left + ScaleLogical(4),
                top,
                Math.Max(0, e.Bounds.Width - ScaleLogical(8)),
                height);
            TextRenderer.DrawText(
                e.Graphics,
                e.SubItem.Text,
                _fileListView.Font,
                bounds,
                foreground,
                alignment
                | TextFormatFlags.VerticalCenter
                | TextFormatFlags.EndEllipsis
                | TextFormatFlags.NoPadding);
        }

        private static void DrawQueueRemoveGlyph(
            Graphics graphics,
            Rectangle bounds,
            Color color)
        {
            if (graphics == null)
            {
                return;
            }
            int centerX = bounds.Left + (bounds.Width / 2);
            int centerY = bounds.Top + (bounds.Height / 2);
            using (var pen = new Pen(color, 1.4f))
            {
                graphics.DrawRectangle(
                    pen,
                    centerX - 4,
                    centerY - 4,
                    8,
                    9);
                graphics.DrawLine(
                    pen,
                    centerX - 6,
                    centerY - 6,
                    centerX + 6,
                    centerY - 6);
                graphics.DrawLine(
                    pen,
                    centerX - 2,
                    centerY - 8,
                    centerX + 2,
                    centerY - 8);
            }
        }

        private static FileLinkQueueRow ResolveQueueRow(
            ListViewItem item)
        {
            return item != null
                ? item.Tag as FileLinkQueueRow
                : null;
        }

        private SelectionUploadState ResolveSelectionState(
            ListViewItem item)
        {
            FileLinkQueueRow row = ResolveQueueRow(item);
            if (row == null || row.Selection == null)
            {
                return null;
            }
            SelectionUploadState state;
            return _selectionStates.TryGetValue(
                row.Selection,
                out state)
                ? state
                : null;
        }

        private void ApplyQueueRowStyle(
            SelectionUploadState state,
            Color backgroundColor,
            Color textColor)
        {
            if (state == null || state.Selection == null)
            {
                return;
            }
            foreach (ListViewItem item in _fileListView.Items)
            {
                FileLinkQueueRow row = ResolveQueueRow(item);
                if (row == null
                    || !ReferenceEquals(row.Selection, state.Selection))
                {
                    continue;
                }
                item.BackColor = backgroundColor;
                item.ForeColor = textColor;
                foreach (ListViewItem.ListViewSubItem subItem
                    in item.SubItems)
                {
                    subItem.BackColor = backgroundColor;
                    if (item.SubItems.IndexOf(subItem) != 2)
                    {
                        subItem.ForeColor = textColor;
                    }
                }
            }
            _fileListView.Invalidate();
        }

        private void SetSelectionQueueStatus(
            SelectionUploadState state,
            string status,
            Color color)
        {
            if (state == null || state.Selection == null)
            {
                return;
            }
            foreach (ListViewItem item in _fileListView.Items)
            {
                FileLinkQueueRow row = ResolveQueueRow(item);
                if (row == null
                    || row.Kind != FileLinkQueueRowKind.File
                    || !ReferenceEquals(row.Selection, state.Selection)
                    || item.SubItems.Count < 3)
                {
                    continue;
                }
                item.SubItems[2].Text = status ?? string.Empty;
                item.SubItems[2].ForeColor = color;
            }
            if (state.Item != null
                && ResolveQueueRow(state.Item) != null
                && ResolveQueueRow(state.Item).Kind
                   == FileLinkQueueRowKind.Folder
                && state.Item.SubItems.Count >= 3)
            {
                state.Item.SubItems[2].Text = status ?? string.Empty;
                state.Item.SubItems[2].ForeColor = color;
            }
            _fileListView.Invalidate();
        }
    }
}
