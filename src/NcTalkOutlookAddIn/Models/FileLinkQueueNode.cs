// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace NcTalkOutlookAddIn.Models
{
    // Immutable file or folder displayed below one queue selection.
    internal sealed class FileLinkQueueNode
    {
        internal FileLinkQueueNode(
            string sourcePath,
            string displayName,
            bool isDirectory,
            long? length,
            DateTime? lastWriteTimeUtc,
            IEnumerable<FileLinkQueueNode> children)
        {
            SourcePath = sourcePath ?? string.Empty;
            DisplayName = displayName ?? string.Empty;
            IsDirectory = isDirectory;
            Length = isDirectory ? null : length;
            LastWriteTimeUtc = isDirectory
                ? null
                : lastWriteTimeUtc;
            Children = new ReadOnlyCollection<FileLinkQueueNode>(
                new List<FileLinkQueueNode>(
                    children ?? Enumerable.Empty<FileLinkQueueNode>()));
        }

        internal string SourcePath { get; private set; }

        internal string DisplayName { get; private set; }

        internal bool IsDirectory { get; private set; }

        internal long? Length { get; private set; }

        internal DateTime? LastWriteTimeUtc { get; private set; }

        internal ReadOnlyCollection<FileLinkQueueNode> Children
        {
            get;
            private set;
        }
    }
}
