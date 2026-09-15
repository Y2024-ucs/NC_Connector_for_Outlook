// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Drawing;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
    internal sealed partial class SettingsForm
    {
        private readonly GroupBox _cardDavSyncGroup = new GroupBox();
        private readonly CheckBox _cardDavEnabledCheckBox = new CheckBox();
        private readonly CheckBox _cardDavCompanyCheckBox = new CheckBox();
        private readonly CheckBox _cardDavPersonalCheckBox = new CheckBox();
        private CardDavSyncPreferences _cardDavPreferences;

        private void InitializeCardDavSettingsSection()
        {
            _cardDavPreferences = CardDavSyncPreferences.Load();

            _cardDavSyncGroup.Text = "Nextcloud-Kontakte";
            _cardDavSyncGroup.Location = new Point(18, 275);
            _cardDavSyncGroup.Size = new Size(440, 132);
            _cardDavSyncGroup.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            _generalTab.Controls.Add(_cardDavSyncGroup);

            _cardDavEnabledCheckBox.Text = "Kontaktsynchronisation aktiv";
            _cardDavEnabledCheckBox.AutoSize = true;
            _cardDavEnabledCheckBox.Location = new Point(12, 25);
            _cardDavEnabledCheckBox.Checked = _cardDavPreferences.Enabled;
            _cardDavEnabledCheckBox.CheckedChanged += delegate { UpdateCardDavSettingsState(); };
            _cardDavSyncGroup.Controls.Add(_cardDavEnabledCheckBox);

            _cardDavCompanyCheckBox.Text = "Firmenverzeichnis synchronisieren";
            _cardDavCompanyCheckBox.AutoSize = true;
            _cardDavCompanyCheckBox.Location = new Point(30, 55);
            _cardDavCompanyCheckBox.Checked = _cardDavPreferences.SyncCompanyDirectory;
            _cardDavSyncGroup.Controls.Add(_cardDavCompanyCheckBox);

            _cardDavPersonalCheckBox.Text = "Persönliche Kontakte synchronisieren";
            _cardDavPersonalCheckBox.AutoSize = true;
            _cardDavPersonalCheckBox.Location = new Point(30, 82);
            _cardDavPersonalCheckBox.Checked = _cardDavPreferences.SyncPersonalContacts;
            _cardDavSyncGroup.Controls.Add(_cardDavPersonalCheckBox);

            var hint = new Label
            {
                Text = "Nur lesend: Nextcloud → Outlook",
                AutoSize = true,
                Location = new Point(30, 108)
            };
            _cardDavSyncGroup.Controls.Add(hint);

            FormClosed += OnCardDavSettingsFormClosed;
            UpdateCardDavSettingsState();
        }

        private void UpdateCardDavSettingsState()
        {
            bool enabled = _cardDavEnabledCheckBox.Checked;
            _cardDavCompanyCheckBox.Enabled = enabled;
            _cardDavPersonalCheckBox.Enabled = enabled;
        }

        private void OnCardDavSettingsFormClosed(object sender, FormClosedEventArgs e)
        {
            if (DialogResult != DialogResult.OK || _cardDavPreferences == null)
            {
                return;
            }

            try
            {
                _cardDavPreferences.Configured = true;
                _cardDavPreferences.Enabled = _cardDavEnabledCheckBox.Checked;
                _cardDavPreferences.SyncCompanyDirectory = _cardDavCompanyCheckBox.Checked;
                _cardDavPreferences.SyncPersonalContacts = _cardDavPersonalCheckBox.Checked;
                _cardDavPreferences.Save();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to save CardDAV settings from settings form.", ex);
            }
        }
    }
}
