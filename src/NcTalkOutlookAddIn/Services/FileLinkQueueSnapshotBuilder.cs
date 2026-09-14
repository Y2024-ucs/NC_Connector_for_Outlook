// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.Services
{
    // Captures the hierarchy shown in the sharing queue without changing either source.
    internal static class FileLinkQueueSnapshotBuilder
    {
        internal static FileLinkQueueNode Build(
            FileLinkSelection selection,
            CancellationToken cancellationToken)
        {
            if (selection == null)
            {
                throw new ArgumentNullException("selection");
            }

            cancellationToken.ThrowIfCancellationRequested();
            return selection.Source == FileLinkSelectionSource.Nextcloud
                ? BuildNextcloud(selection, cancellationToken)
                : BuildLocal(selection, cancellationToken);
        }

        private static FileLinkQueueNode BuildLocal(
            FileLinkSelection selection,
            CancellationToken cancellationToken)
        {
            if (selection.SelectionType == FileLinkSelectionType.File)
            {
                return BuildLocalFile(
                    selection.LocalPath,
                    cancellationToken);
            }

            return BuildLocalDirectory(
                selection.LocalPath,
                cancellationToken);
        }

        private static FileLinkQueueNode BuildLocalDirectory(
            string path,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var directory = new DirectoryInfo(path);
            EnsureNotLinked(directory);

            var children = new List<FileLinkQueueNode>();
            foreach (DirectoryInfo child in directory
                .GetDirectories()
                .OrderBy(
                    item => item.Name,
                    StringComparer.CurrentCultureIgnoreCase))
            {
                children.Add(BuildLocalDirectory(
                    child.FullName,
                    cancellationToken));
            }
            foreach (FileInfo child in directory
                .GetFiles()
                .OrderBy(
                    item => item.Name,
                    StringComparer.CurrentCultureIgnoreCase))
            {
                children.Add(BuildLocalFile(
                    child.FullName,
                    cancellationToken));
            }

            string displayName = directory.Name;
            if (string.IsNullOrWhiteSpace(displayName))
            {
                displayName = directory.FullName;
            }
            return new FileLinkQueueNode(
                directory.FullName,
                displayName,
                true,
                null,
                null,
                children);
        }

        private static FileLinkQueueNode BuildLocalFile(
            string path,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var file = new FileInfo(path);
            EnsureNotLinked(file);
            return new FileLinkQueueNode(
                file.FullName,
                file.Name,
                false,
                file.Length,
                file.LastWriteTimeUtc,
                null);
        }

        private static void EnsureNotLinked(FileSystemInfo item)
        {
            if (item == null
                || (item.Attributes & FileAttributes.ReparsePoint) == 0)
            {
                return;
            }
            throw new IOException(
                Strings.FileLinkUploadLinkedItemUnsupported);
        }

        private static FileLinkQueueNode BuildNextcloud(
            FileLinkSelection selection,
            CancellationToken cancellationToken)
        {
            IList<NextcloudStorageEntry> entries =
                selection.NextcloudEntries;
            if (entries == null || entries.Count == 0)
            {
                throw CreateChangedSourceException();
            }

            if (selection.SelectionType == FileLinkSelectionType.File)
            {
                NextcloudStorageEntry file = entries.FirstOrDefault(
                    item => item != null
                            && !item.IsDirectory
                            && string.Equals(
                                item.RelativePath,
                                selection.NextcloudPath,
                                StringComparison.Ordinal));
                if (file == null)
                {
                    throw CreateChangedSourceException();
                }
                return new FileLinkQueueNode(
                    file.RelativePath,
                    file.DisplayName,
                    false,
                    file.Length,
                    file.LastModifiedUtc,
                    null);
            }

            NextcloudStorageEntry rootEntry = entries.FirstOrDefault(
                item => item != null
                        && item.IsDirectory
                        && string.Equals(
                            item.RelativePath,
                            selection.NextcloudPath,
                            StringComparison.Ordinal));
            if (rootEntry == null)
            {
                throw CreateChangedSourceException();
            }

            var root = new MutableNode(
                rootEntry.RelativePath,
                selection.DisplayName,
                true,
                null);
            var nodes = new Dictionary<string, MutableNode>(
                StringComparer.Ordinal)
            {
                { rootEntry.RelativePath, root }
            };
            IEnumerable<NextcloudStorageEntry> descendants = entries
                .Where(
                    item => item != null
                            && !string.Equals(
                                item.RelativePath,
                                rootEntry.RelativePath,
                                StringComparison.Ordinal))
                .OrderBy(item => NextcloudPath.GetDepth(item.RelativePath))
                .ThenBy(item => item.IsDirectory ? 0 : 1)
                .ThenBy(
                    item => item.DisplayName,
                    StringComparer.CurrentCultureIgnoreCase);
            foreach (NextcloudStorageEntry entry in descendants)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!IsDescendant(
                    rootEntry.RelativePath,
                    entry.RelativePath))
                {
                    throw CreateChangedSourceException();
                }

                string parentPath = NextcloudPath.GetParent(
                    entry.RelativePath);
                MutableNode parent;
                if (!nodes.TryGetValue(parentPath, out parent)
                    || nodes.ContainsKey(entry.RelativePath))
                {
                    throw CreateChangedSourceException();
                }

                var node = new MutableNode(
                    entry.RelativePath,
                    entry.DisplayName,
                    entry.IsDirectory,
                    entry.IsDirectory ? null : (long?)entry.Length,
                    entry.IsDirectory
                        ? null
                        : entry.LastModifiedUtc);
                nodes.Add(entry.RelativePath, node);
                parent.Children.Add(node);
            }

            return ConvertNode(root);
        }

        private static bool IsDescendant(
            string parent,
            string candidate)
        {
            string normalizedParent = NextcloudPath.Normalize(parent);
            string normalizedCandidate = NextcloudPath.Normalize(candidate);
            if (normalizedParent.Length == 0)
            {
                return normalizedCandidate.Length > 0;
            }
            return normalizedCandidate.StartsWith(
                normalizedParent + "/",
                StringComparison.Ordinal);
        }

        private static FileLinkQueueNode ConvertNode(MutableNode source)
        {
            IEnumerable<MutableNode> orderedChildren = source.Children
                .OrderBy(item => item.IsDirectory ? 0 : 1)
                .ThenBy(
                    item => item.DisplayName,
                    StringComparer.CurrentCultureIgnoreCase);
            return new FileLinkQueueNode(
                source.SourcePath,
                source.DisplayName,
                source.IsDirectory,
                source.Length,
                source.LastWriteTimeUtc,
                orderedChildren.Select(ConvertNode));
        }

        private static IOException CreateChangedSourceException()
        {
            return new IOException(
                Strings.FileLinkUploadSourceChanged);
        }

        private sealed class MutableNode
        {
            internal MutableNode(
                string sourcePath,
                string displayName,
                bool isDirectory,
                long? length,
                DateTime? lastWriteTimeUtc = null)
            {
                SourcePath = sourcePath ?? string.Empty;
                DisplayName = displayName ?? string.Empty;
                IsDirectory = isDirectory;
                Length = length;
                LastWriteTimeUtc = lastWriteTimeUtc;
                Children = new List<MutableNode>();
            }

            internal string SourcePath { get; private set; }

            internal string DisplayName { get; private set; }

            internal bool IsDirectory { get; private set; }

            internal long? Length { get; private set; }

            internal DateTime? LastWriteTimeUtc { get; private set; }

            internal List<MutableNode> Children { get; private set; }
        }
    }
}
