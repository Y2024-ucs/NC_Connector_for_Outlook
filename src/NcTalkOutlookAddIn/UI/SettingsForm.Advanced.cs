// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Net;
using System.Reflection;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Settings;
using NcTalkOutlookAddIn.Utilities;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NcTalkOutlookAddIn.UI
{
    internal sealed partial class SettingsForm
    {
        private void ApplyAdvancedTabLayout()
        {
            int left = ScaleLogical(24);
            int labelToComboGap = ScaleLogical(16);
            int comboLeft = left + Math.Max(_ifbCacheHoursLabel.PreferredSize.Width, Math.Max(_shareBlockLangLabel.PreferredSize.Width, _eventDescriptionLangLabel.PreferredSize.Width)) + labelToComboGap;
            int rightMargin = ScaleLogical(24);
            int comboWidth = Math.Max(ScaleLogical(160), _advancedTab.ClientSize.Width - comboLeft - rightMargin);
            int rowTop = ScaleLogical(24);
            int rowGap = ScaleLogical(34);

            int ifbComboHeight = Math.Max(_ifbCacheHoursCombo.Height, _ifbCacheHoursCombo.PreferredHeight + ScaleLogical(2));
            _ifbCacheHoursLabel.Location = new Point(left, rowTop);
            _ifbCacheHoursCombo.SetBounds(comboLeft, rowTop - ScaleLogical(2), Math.Max(ScaleLogical(90), _ifbCacheHoursCombo.Width), ifbComboHeight);

            int shareLabelTop = rowTop + rowGap + ScaleLogical(12);
            int shareComboHeight = Math.Max(_shareBlockLangCombo.Height, _shareBlockLangCombo.PreferredHeight + ScaleLogical(2));
            _shareBlockLangLabel.Location = new Point(left, shareLabelTop);
            _shareBlockLangCombo.SetBounds(comboLeft, shareLabelTop - ScaleLogical(2), comboWidth, shareComboHeight);

            int eventLabelTop = _shareBlockLangLabel.Bottom + rowGap;
            int eventComboHeight = Math.Max(_eventDescriptionLangCombo.Height, _eventDescriptionLangCombo.PreferredHeight + ScaleLogical(2));
            _eventDescriptionLangLabel.Location = new Point(left, eventLabelTop);
            _eventDescriptionLangCombo.SetBounds(comboLeft, eventLabelTop - ScaleLogical(2), comboWidth, eventComboHeight);

            int groupTop = _eventDescriptionLangCombo.Bottom + ScaleLogical(14);
            int groupWidth = Math.Max(ScaleLogical(320), _advancedTab.ClientSize.Width - left - rightMargin);
            int tlsTop = groupTop;
            _tlsSettingsGroup.SetBounds(left, tlsTop, groupWidth, ScaleLogical(134));
            _tlsHintLabel.MaximumSize = new Size(Math.Max(ScaleLogical(200), _tlsSettingsGroup.ClientSize.Width - ScaleLogical(20)), 0);
            _tlsHintLabel.AutoSize = true;
        }

        private void InitializeAdvancedTab()
        {
            _advancedTab.AutoScroll = true;
            _advancedTab.Padding = new Padding(12);

            _ifbCacheHoursLabel.Text = Strings.LabelIfbCacheHours;
            _ifbCacheHoursLabel.Location = new Point(24, 24);
            _ifbCacheHoursLabel.AutoSize = true;
            _advancedTab.Controls.Add(_ifbCacheHoursLabel);

            _ifbCacheHoursCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _ifbCacheHoursCombo.IntegralHeight = false;
            _ifbCacheHoursCombo.Location = new Point(260, 22);
            _ifbCacheHoursCombo.Width = 80;
            _ifbCacheHoursCombo.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            for (int i = 1; i <= 24; i++)
            {
                _ifbCacheHoursCombo.Items.Add(i.ToString());
            }
            _advancedTab.Controls.Add(_ifbCacheHoursCombo);

            int langTop = 70;

            _shareBlockLangLabel.Text = Strings.AdvancedShareBlockLangLabel;
            _shareBlockLangLabel.Location = new Point(24, langTop);
            _shareBlockLangLabel.AutoSize = true;
            _advancedTab.Controls.Add(_shareBlockLangLabel);

            _shareBlockLangCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _shareBlockLangCombo.DrawMode = DrawMode.OwnerDrawFixed;
            _shareBlockLangCombo.IntegralHeight = false;
            _shareBlockLangCombo.Location = new Point(260, langTop - 2);
            _shareBlockLangCombo.Width = 240;
            _shareBlockLangCombo.DrawItem += HandleLanguageComboDrawItem;
            _shareBlockLangCombo.SelectionChangeCommitted += HandleLanguageComboSelectionCommitted;
            PopulateLanguageOverrideCombo(_shareBlockLangCombo, "share");
            _advancedTab.Controls.Add(_shareBlockLangCombo);

            _eventDescriptionLangLabel.Text = Strings.AdvancedEventDescriptionLangLabel;
            _eventDescriptionLangLabel.Location = new Point(24, langTop + 32);
            _eventDescriptionLangLabel.AutoSize = true;
            _advancedTab.Controls.Add(_eventDescriptionLangLabel);

            _eventDescriptionLangCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _eventDescriptionLangCombo.DrawMode = DrawMode.OwnerDrawFixed;
            _eventDescriptionLangCombo.IntegralHeight = false;
            _eventDescriptionLangCombo.Location = new Point(260, langTop + 30);
            _eventDescriptionLangCombo.Width = 240;
            _eventDescriptionLangCombo.DrawItem += HandleLanguageComboDrawItem;
            _eventDescriptionLangCombo.SelectionChangeCommitted += HandleLanguageComboSelectionCommitted;
            PopulateLanguageOverrideCombo(_eventDescriptionLangCombo, "talk");
            _advancedTab.Controls.Add(_eventDescriptionLangCombo);

            _tlsSettingsGroup.Text = Strings.AdvancedTlsHeading;
            _tlsSettingsGroup.Location = new Point(24, langTop + 72);
            _tlsSettingsGroup.Size = new Size(520, 132);
            _advancedTab.Controls.Add(_tlsSettingsGroup);

            _tlsUseSystemDefaultCheckBox.Text = Strings.AdvancedTlsUseSystemDefaultLabel;
            _tlsUseSystemDefaultCheckBox.AutoSize = true;
            _tlsUseSystemDefaultCheckBox.Location = new Point(12, 24);
            _tlsUseSystemDefaultCheckBox.CheckedChanged += OnTlsSelectionChanged;
            _tlsSettingsGroup.Controls.Add(_tlsUseSystemDefaultCheckBox);

            _tlsEnable12CheckBox.Text = Strings.AdvancedTlsEnable12Label;
            _tlsEnable12CheckBox.AutoSize = true;
            _tlsEnable12CheckBox.Location = new Point(12, 48);
            _tlsEnable12CheckBox.CheckedChanged += OnTlsSelectionChanged;
            _tlsSettingsGroup.Controls.Add(_tlsEnable12CheckBox);

            _tlsEnable13CheckBox.Text = Strings.AdvancedTlsEnable13Label;
            _tlsEnable13CheckBox.AutoSize = true;
            _tlsEnable13CheckBox.Location = new Point(12, 72);
            _tlsEnable13CheckBox.CheckedChanged += OnTlsSelectionChanged;
            _tlsSettingsGroup.Controls.Add(_tlsEnable13CheckBox);

            _tlsHintLabel.Text = Strings.AdvancedTlsHint;
            _tlsHintLabel.AutoSize = true;
            _tlsHintLabel.MaximumSize = new Size(492, 0);
            _tlsHintLabel.Location = new Point(12, 94);
            _tlsHintLabel.ForeColor = Color.DimGray;
            _tlsSettingsGroup.Controls.Add(_tlsHintLabel);
        }

        private void UpdateTlsOptionsState()
        {
            bool useSystemDefault = _tlsUseSystemDefaultCheckBox.Checked;
            bool allowCustom = !useSystemDefault && !_isBusy;

            _tlsEnable12CheckBox.Enabled = allowCustom;
            _tlsEnable13CheckBox.Enabled = allowCustom;
        }

        private void OnTlsSelectionChanged(object sender, EventArgs e)
        {
            UpdateTlsOptionsState();
            ApplyTlsRuntimePreview("settings_tls_changed");
        }

        private void ApplyTlsRuntimePreview(string source)
        {
            if (_suppressImmediateTlsApply)
            {
                return;
            }
            if (!_tlsUseSystemDefaultCheckBox.Checked
                && !_tlsEnable12CheckBox.Checked
                && !_tlsEnable13CheckBox.Checked)
            {
                SetStatus(Strings.TransportTlsSelectionRequired, true);
                return;
            }
            try
            {
                TransportSecurityConfigurator.Apply(
                    _tlsUseSystemDefaultCheckBox.Checked,
                    _tlsEnable12CheckBox.Checked,
                    _tlsEnable13CheckBox.Checked,
                    source);
                SetStatus(string.Empty, false);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.Core,
                    "Failed to apply live TLS runtime preview from settings UI.",
                    ex);
                SetStatus(string.Format(Strings.TransportTlsApplyFailed, ex.Message), true);
            }
        }

    }
}
