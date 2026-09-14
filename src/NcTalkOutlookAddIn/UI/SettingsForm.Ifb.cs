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
        private void ApplyIfbTabLayout()
        {
            int left = ScaleLogical(24);
            int top = ScaleLogical(20);
            int rowGap = ScaleLogical(18);
            int labelToComboGap = ScaleLogical(14);

            _ifbEnabledCheckBox.Location = new Point(left, top);

            int daysLabelTop = _ifbEnabledCheckBox.Bottom + rowGap;
            _ifbDaysLabel.Location = new Point(left, daysLabelTop);

            int portLabelTop = _ifbDaysLabel.Bottom + rowGap;
            _ifbPortLabel.Location = new Point(left, portLabelTop);

            int comboLeft = Math.Max(_ifbDaysLabel.Right, _ifbPortLabel.Right) + labelToComboGap;
            int comboHeight = Math.Max(_ifbDaysCombo.Height, _ifbDaysCombo.PreferredHeight + ScaleLogical(2));
            _ifbDaysCombo.SetBounds(comboLeft, daysLabelTop - ScaleLogical(2), Math.Max(ScaleLogical(90), _ifbDaysCombo.Width), comboHeight);

            int portHeight = Math.Max(_ifbPortUpDown.Height, _ifbPortUpDown.PreferredHeight + ScaleLogical(2));
            _ifbPortUpDown.SetBounds(comboLeft, portLabelTop - ScaleLogical(2), Math.Max(ScaleLogical(110), _ifbPortUpDown.Width), portHeight);
        }

        private void InitializeIfbTab()
        {
            _ifbTab.AutoScroll = true;
            _ifbTab.Padding = new Padding(12);

            _ifbEnabledCheckBox.Text = Strings.CheckIfbEnabled;
            _ifbEnabledCheckBox.AutoSize = true;
            _ifbEnabledCheckBox.Location = new Point(24, 20);
            _ifbEnabledCheckBox.CheckedChanged += OnIfbEnabledChanged;
            _ifbTab.Controls.Add(_ifbEnabledCheckBox);

            _ifbDaysLabel.Text = Strings.LabelIfbDays;
            _ifbDaysLabel.Location = new Point(24, 60);
            _ifbDaysLabel.AutoSize = true;
            _ifbTab.Controls.Add(_ifbDaysLabel);

            _ifbDaysCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _ifbDaysCombo.IntegralHeight = false;
            _ifbDaysCombo.Location = new Point(200, 58);
            _ifbDaysCombo.Width = 100;
            _ifbDaysCombo.Items.AddRange(new object[] { "10", "30", "60", "90" });
            _ifbTab.Controls.Add(_ifbDaysCombo);

            _ifbPortLabel.Text = Strings.LabelIfbPort;
            _ifbPortLabel.Location = new Point(24, 94);
            _ifbPortLabel.AutoSize = true;
            _ifbTab.Controls.Add(_ifbPortLabel);

            _ifbPortUpDown.Location = new Point(200, 92);
            _ifbPortUpDown.Width = 110;
            _ifbPortUpDown.Minimum = AddinSettings.MinIfbPort;
            _ifbPortUpDown.Maximum = AddinSettings.MaxIfbPort;
            _ifbPortUpDown.Value = AddinSettings.DefaultIfbPort;
            _ifbTab.Controls.Add(_ifbPortUpDown);
        }

        private void OnIfbEnabledChanged(object sender, EventArgs e)
        {
            _ifbDefaultApplied = true;

            if (_isBusy)
            {
                return;
            }

            UpdateControlState();
        }

    }
}
