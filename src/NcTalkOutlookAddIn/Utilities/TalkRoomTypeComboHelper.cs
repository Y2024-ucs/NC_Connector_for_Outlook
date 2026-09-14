// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;

namespace NcTalkOutlookAddIn.Utilities
{
    internal static class TalkRoomTypeComboHelper
    {
        internal static void Populate(
            ComboBox comboBox,
            string eventConversationLabel,
            string standardRoomLabel)
        {
            if (comboBox == null)
            {
                return;
            }

            comboBox.Items.Clear();
            comboBox.Items.Add(
                new TalkRoomTypeOption(
                    TalkRoomType.EventConversation,
                    eventConversationLabel));
            comboBox.Items.Add(
                new TalkRoomTypeOption(
                    TalkRoomType.StandardRoom,
                    standardRoomLabel));
        }

        internal static void Select(ComboBox comboBox, TalkRoomType value)
        {
            if (comboBox == null)
            {
                return;
            }

            foreach (object item in comboBox.Items)
            {
                var option = item as TalkRoomTypeOption;
                if (option != null && option.Value == value)
                {
                    comboBox.SelectedItem = option;
                    return;
                }
            }

            if (comboBox.Items.Count > 0)
            {
                comboBox.SelectedIndex = 0;
            }
        }

        internal static TalkRoomType GetSelected(
            ComboBox comboBox,
            TalkRoomType fallback)
        {
            var selected = comboBox != null
                ? comboBox.SelectedItem as TalkRoomTypeOption
                : null;
            return selected != null ? selected.Value : fallback;
        }

        private sealed class TalkRoomTypeOption
        {
            internal TalkRoomTypeOption(TalkRoomType value, string label)
            {
                Value = value;
                Label = label ?? value.ToString();
            }

            internal TalkRoomType Value { get; private set; }

            private string Label { get; set; }

            public override string ToString()
            {
                return Label;
            }
        }
    }
}
