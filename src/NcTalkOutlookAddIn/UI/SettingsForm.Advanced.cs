// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Net;
using System.Reflection;
using System.Threading.Tasks;
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
            _updateSettingsGroup.SetBounds(left, groupTop, groupWidth, ScaleLogical(286));
            int updateInnerWidth = Math.Max(ScaleLogical(220), _updateSettingsGroup.ClientSize.Width - ScaleLogical(24));
            _updateNotifyCheckBox.Location = new Point(ScaleLogical(12), ScaleLogical(24));
            _updateInstalledVersionLabel.SetBounds(ScaleLogical(12), ScaleLogical(54), updateInnerWidth, ScaleLogical(20));
            _updateLatestVersionLabel.SetBounds(ScaleLogical(12), ScaleLogical(78), updateInnerWidth, ScaleLogical(20));
            _updateLastCheckedLabel.SetBounds(ScaleLogical(12), ScaleLogical(102), updateInnerWidth, ScaleLogical(20));
            _updateCheckButton.SetBounds(ScaleLogical(12), ScaleLogical(130), ScaleLogical(120), ScaleLogical(28));
            _updateDownloadLink.SetBounds(_updateCheckButton.Right + ScaleLogical(16), _updateCheckButton.Top + ScaleLogical(6), Math.Max(ScaleLogical(140), updateInnerWidth - _updateCheckButton.Width - ScaleLogical(28)), ScaleLogical(22));
            _updateChangelogLabel.Location = new Point(ScaleLogical(12), ScaleLogical(166));
            _updateChangelogTextBox.SetBounds(ScaleLogical(12), ScaleLogical(186), updateInnerWidth, ScaleLogical(86));

            int tlsTop = _updateSettingsGroup.Bottom + ScaleLogical(14);
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

            _updateSettingsGroup.Text = "Updates (IBP-Abspaltung – deaktiviert)";
            _updateSettingsGroup.Location = new Point(24, langTop + 72);
            _updateSettingsGroup.Size = new Size(520, 286);
            _advancedTab.Controls.Add(_updateSettingsGroup);

            _updateNotifyCheckBox.Text = "Offizielle Update-Benachrichtigungen deaktiviert";
            _updateNotifyCheckBox.AutoSize = true;
            _updateNotifyCheckBox.Location = new Point(12, 24);
            _updateNotifyCheckBox.Checked = false;
            _updateNotifyCheckBox.Enabled = false;
            _updateSettingsGroup.Controls.Add(_updateNotifyCheckBox);

            _updateInstalledVersionLabel.AutoSize = false;
            _updateInstalledVersionLabel.Location = new Point(12, 54);
            _updateInstalledVersionLabel.Size = new Size(480, 20);
            _updateSettingsGroup.Controls.Add(_updateInstalledVersionLabel);

            _updateLatestVersionLabel.AutoSize = false;
            _updateLatestVersionLabel.Location = new Point(12, 78);
            _updateLatestVersionLabel.Size = new Size(480, 20);
            _updateSettingsGroup.Controls.Add(_updateLatestVersionLabel);

            _updateLastCheckedLabel.AutoSize = false;
            _updateLastCheckedLabel.Location = new Point(12, 102);
            _updateLastCheckedLabel.Size = new Size(480, 20);
            _updateSettingsGroup.Controls.Add(_updateLastCheckedLabel);

            _updateCheckButton.Text = Strings.UpdateCheckNowButton;
            _updateCheckButton.Location = new Point(12, 130);
            _updateCheckButton.Size = new Size(120, 28);
            _updateCheckButton.Enabled = false;
            _updateSettingsGroup.Controls.Add(_updateCheckButton);

            _updateDownloadLink.Text = "Upstream-Download deaktiviert";
            _updateDownloadLink.Location = new Point(148, 136);
            _updateDownloadLink.AutoSize = false;
            _updateDownloadLink.Size = new Size(320, 22);
            _updateDownloadLink.Enabled = false;
            _updateSettingsGroup.Controls.Add(_updateDownloadLink);

            _updateChangelogLabel.Text = Strings.UpdateChangelogHeading;
            _updateChangelogLabel.Location = new Point(12, 166);
            _updateChangelogLabel.AutoSize = true;
            _updateSettingsGroup.Controls.Add(_updateChangelogLabel);

            _updateChangelogTextBox.ReadOnly = true;
            _updateChangelogTextBox.Multiline = true;
            _updateChangelogTextBox.ScrollBars = ScrollBars.Vertical;
            _updateChangelogTextBox.Location = new Point(12, 186);
            _updateChangelogTextBox.Size = new Size(480, 86);
            _updateChangelogTextBox.Text = "Die offizielle Update-Funktion ist in der IBP-Abspaltung deaktiviert, damit keine Upstream-Version die IBP-Anpassungen überschreibt.";
            _updateSettingsGroup.Controls.Add(_updateChangelogTextBox);

            _tlsSettingsGroup.Text = Strings.AdvancedTlsHeading;
            _tlsSettingsGroup.Location = new Point(24, langTop + 372);
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

        private void UpdateUpdateCheckSection()
        {
            string installedVersion = AddinVersionInfo.GetVersion();
            if (string.IsNullOrWhiteSpace(installedVersion))
            {
                installedVersion = Strings.TalkVersionUnknown;
            }

            _updateNotifyCheckBox.Checked = false;
            _updateNotifyCheckBox.Enabled = false;
            _updateInstalledVersionLabel.Text = string.Format(Strings.UpdateInstalledVersionFormat, installedVersion);
            _updateLatestVersionLabel.Text = "Neueste Version: Upstream-Prüfung deaktiviert";
            _updateLastCheckedLabel.Text = "Letzte Prüfung: deaktiviert";
            _updateOpenUrl = string.Empty;
            _updateDownloadLink.Text = "Upstream-Download deaktiviert";
            _updateDownloadLink.Enabled = false;
            _updateCheckButton.Enabled = false;
            _updateChangelogTextBox.Text = "Die offizielle Update-Funktion ist in der IBP-Abspaltung deaktiviert, damit keine Upstream-Version die IBP-Anpassungen überschreibt.";
        }

        private static string FormatUpdateCheckedAt(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return Strings.UpdateNotChecked;
            }

            DateTime parsed;
            if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out parsed))
            {
                return parsed.ToLocalTime().ToString("g", CultureInfo.CurrentCulture);
            }

            return Strings.UpdateNotChecked;
        }

        private async void OnUpdateCheckButtonClick(object sender, EventArgs e)
        {
            await Task.Yield();
            SetStatus("Offizielle Updates sind in der IBP-Abspaltung deaktiviert.", false);
        }

        private void OnUpdateDownloadLinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SetStatus("Offizielle Updates sind in der IBP-Abspaltung deaktiviert.", false);
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
