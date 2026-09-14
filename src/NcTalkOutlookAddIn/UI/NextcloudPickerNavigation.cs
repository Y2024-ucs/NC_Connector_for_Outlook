// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
    internal sealed class NextcloudPickerNavigation
    {
        private readonly List<string> _history = new List<string>();
        private int _index = -1;

        internal bool CanGoBack { get { return _index > 0; } }
        internal bool CanGoForward { get { return _index >= 0 && _index < _history.Count - 1; } }

        internal bool TryGetTarget(int offset, out int targetIndex, out string targetPath)
        {
            targetIndex = _index + offset;
            targetPath = string.Empty;
            if (targetIndex < 0 || targetIndex >= _history.Count)
            {
                return false;
            }
            targetPath = _history[targetIndex];
            return true;
        }

        internal void CompleteHistoryNavigation(int targetIndex, string loadedPath)
        {
            _index = targetIndex;
            _history[targetIndex] = loadedPath;
        }

        internal void Record(string relativePath)
        {
            string normalizedPath = NextcloudPath.Normalize(relativePath);
            if (_index >= 0
                && _index < _history.Count
                && string.Equals(_history[_index], normalizedPath, StringComparison.Ordinal))
            {
                return;
            }
            int firstForwardIndex = _index + 1;
            if (firstForwardIndex < _history.Count)
            {
                _history.RemoveRange(firstForwardIndex, _history.Count - firstForwardIndex);
            }
            _history.Add(normalizedPath);
            _index = _history.Count - 1;
        }
    }
}
