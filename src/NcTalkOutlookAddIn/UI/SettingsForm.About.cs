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
        private void ApplyAboutTabLayout()
        {
            int left = ScaleLogical(18);
            int right = ScaleLogical(18);
            int top = ScaleLogical(20);
            int gap = ScaleLogical(10);
            int contentWidth = Math.Max(ScaleLogical(220), _aboutTab.ClientSize.Width - left - right);

            _aboutVersionLabel.Location = new Point(left, top);
            _aboutCopyrightLabel.Location = new Point(left, _aboutVersionLabel.Bottom + gap);

            _aboutLicenseLabel.Location = new Point(left, _aboutCopyrightLabel.Bottom + gap);
            _aboutLicenseLink.Location = new Point(_aboutLicenseLabel.Right + ScaleLogical(8), _aboutLicenseLabel.Top);

            _aboutHomepageLabel.Location = new Point(left, _aboutLicenseLabel.Bottom + gap);
            _aboutHomepageLink.Location = new Point(_aboutHomepageLabel.Right + ScaleLogical(8), _aboutHomepageLabel.Top);

            _aboutOverviewLabel.Location = new Point(left, _aboutHomepageLabel.Bottom + ScaleLogical(16));
            _aboutOverviewLabel.MaximumSize = new Size(contentWidth, 0);
            _aboutOverviewLabel.AutoSize = true;

            _aboutMoreInfoLabel.Location = new Point(left, _aboutOverviewLabel.Bottom + gap);
            _aboutMoreInfoLink.Location = new Point(left, _aboutMoreInfoLabel.Bottom + ScaleLogical(6));

            _aboutSupportNoteLabel.Location = new Point(left, _aboutMoreInfoLink.Bottom + ScaleLogical(16));
            _aboutSupportNoteLabel.MaximumSize = new Size(contentWidth, 0);
            _aboutSupportNoteLabel.AutoSize = true;

            _aboutSupportHeadingLabel.Location = new Point(left, _aboutSupportNoteLabel.Bottom + gap);
            _aboutSupportLink.Location = new Point(left, _aboutSupportHeadingLabel.Bottom + ScaleLogical(8));
        }

        private void InitializeAboutTab()
        {
            _aboutTab.Padding = new Padding(12);
            _aboutTab.AutoScroll = true;

            _aboutVersionLabel.AutoSize = true;
            _aboutTab.Controls.Add(_aboutVersionLabel);

            _aboutCopyrightLabel.AutoSize = true;
            _aboutTab.Controls.Add(_aboutCopyrightLabel);

            _aboutLicenseLabel.Text = Strings.AboutLicenseLabel;
            _aboutLicenseLabel.AutoSize = true;
            _aboutTab.Controls.Add(_aboutLicenseLabel);

            _aboutLicenseLink.AutoSize = true;
            _aboutLicenseLink.Text = Strings.AboutLicenseLink;
            _aboutLicenseLink.Padding = new Padding(1, 0, 0, 0);
            _aboutLicenseLink.LinkClicked += OnAboutLicenseLinkClicked;
            _aboutTab.Controls.Add(_aboutLicenseLink);

            _aboutHomepageLabel.Text = Strings.AboutHomepageLabel;
            _aboutHomepageLabel.AutoSize = true;
            _aboutTab.Controls.Add(_aboutHomepageLabel);

            _aboutHomepageLink.Text = Strings.AboutHomepageLink;
            _aboutHomepageLink.AutoSize = true;
            _aboutHomepageLink.LinkClicked += (sender, args) =>
                BrowserLauncher.OpenUrl(
                    NcConnectorHomepageUrl,
                    LogCategories.Core,
                    "Failed to open NC Connector homepage URL.");
            _aboutTab.Controls.Add(_aboutHomepageLink);

            _aboutOverviewLabel.Text = Strings.AboutOverviewText;
            _aboutOverviewLabel.AutoSize = true;
            _aboutOverviewLabel.MaximumSize = new Size(Math.Max(ScaleLogical(220), _aboutTab.ClientSize.Width - ScaleLogical(36)), 0);
            _aboutTab.Controls.Add(_aboutOverviewLabel);

            _aboutMoreInfoLabel.Text = Strings.AboutMoreInfoLabel;
            _aboutMoreInfoLabel.AutoSize = true;
            _aboutMoreInfoLabel.Font = new Font(_aboutMoreInfoLabel.Font, FontStyle.Bold);
            _aboutTab.Controls.Add(_aboutMoreInfoLabel);

            _aboutMoreInfoLink.Text = Strings.AboutMoreInfoLink;
            _aboutMoreInfoLink.AutoSize = true;
            _aboutMoreInfoLink.LinkClicked += (sender, args) =>
                BrowserLauncher.OpenUrl(
                    NcConnectorHomepageUrl,
                    LogCategories.Core,
                    "Failed to open NC Connector information URL.");
            _aboutTab.Controls.Add(_aboutMoreInfoLink);

            _aboutSupportNoteLabel.Text = Strings.AboutSupportNote;
            _aboutSupportNoteLabel.AutoSize = true;
            _aboutSupportNoteLabel.ForeColor = Color.DimGray;
            _aboutTab.Controls.Add(_aboutSupportNoteLabel);

            _aboutSupportHeadingLabel.Text = Strings.AboutSupportHeading;
            _aboutSupportHeadingLabel.AutoSize = true;
            _aboutSupportHeadingLabel.Font = new Font(_aboutSupportHeadingLabel.Font, FontStyle.Bold);
            _aboutTab.Controls.Add(_aboutSupportHeadingLabel);

            _aboutSupportLink.Text = Strings.AboutSupportLink;
            _aboutSupportLink.AutoSize = true;
            _aboutSupportLink.LinkClicked += (sender, args) =>
                BrowserLauncher.OpenUrl(
                    "https://www.paypal.com/donate/?hosted_button_id=FTZWNRNKVKUN6",
                    LogCategories.Core,
                    "Failed to open support donation URL.");
            _aboutTab.Controls.Add(_aboutSupportLink);

            ApplyAboutTabLayout();
        }

        private void UpdateAboutTab()
        {
            string version = AddinVersionInfo.GetVersion();
            if (string.IsNullOrWhiteSpace(version))
            {
                version = "n/a";
            }
            _aboutVersionLabel.Text = string.Format(Strings.AboutVersionFormat, version);
            _aboutCopyrightLabel.Text = Strings.AboutCopyright;
        }

        private void OnAboutLicenseLinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string licensePath = ResolveLicensePath();
            if (File.Exists(licensePath))
            {
                string openError;
                bool opened = BrowserLauncher.OpenTarget(
                    licensePath,
                    LogCategories.Core,
                    "Failed to open license file.",
                    out openError);
                if (!opened)
                {
                    MessageBox.Show(
                        string.Format(Strings.LicenseFileOpenErrorMessage, openError ?? string.Empty),
                        Strings.DialogTitle,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
                return;
            }

            MessageBox.Show(
                string.Format(Strings.LicenseFileMissingMessage, licensePath),
                Strings.DialogTitle,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private static string ResolveLicensePath()
        {
            try
            {
                string assemblyPath = Assembly.GetExecutingAssembly().Location;
                if (!string.IsNullOrEmpty(assemblyPath))
                {
                    string baseDir = Path.GetDirectoryName(assemblyPath) ?? string.Empty;
                    return Path.Combine(baseDir, "License.txt");
                }
                string appBase = AppDomain.CurrentDomain.BaseDirectory ?? string.Empty;
                return Path.Combine(appBase, "License.txt");
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to resolve license path.", ex);
                return "License.txt";
            }
        }

    }
}
