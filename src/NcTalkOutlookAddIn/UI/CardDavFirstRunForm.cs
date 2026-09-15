// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System.Drawing;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
    internal sealed class CardDavFirstRunForm : ScaledForm
    {
        private readonly CheckBox _enabledCheckBox = new CheckBox();
        private readonly CheckBox _companyCheckBox = new CheckBox();
        private readonly CheckBox _personalCheckBox = new CheckBox();

        internal CardDavFirstRunForm(CardDavSyncPreferences preferences)
        {
            Text = "Nextcloud-Kontakte synchronisieren";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(520, 250);
            Icon = BrandingAssets.GetAppIcon(32);

            var title = new Label
            {
                Text = "Welche Nextcloud-Kontakte sollen in Outlook angezeigt werden?",
                AutoSize = false,
                Font = new Font(Font, FontStyle.Bold),
                Location = new Point(20, 20),
                Size = new Size(480, 36)
            };
            Controls.Add(title);

            var hint = new Label
            {
                Text = "Die Synchronisation ist nur lesend: Nextcloud → Outlook. Änderungen in Outlook werden nicht zurückgeschrieben.",
                AutoSize = false,
                Location = new Point(20, 58),
                Size = new Size(480, 44)
            };
            Controls.Add(hint);

            _enabledCheckBox.Text = "Kontaktsynchronisation aktivieren";
            _enabledCheckBox.AutoSize = true;
            _enabledCheckBox.Location = new Point(20, 110);
            _enabledCheckBox.Checked = preferences == null || preferences.Enabled;
            _enabledCheckBox.CheckedChanged += delegate { UpdateState(); };
            Controls.Add(_enabledCheckBox);

            _companyCheckBox.Text = "Firmenverzeichnis synchronisieren";
            _companyCheckBox.AutoSize = true;
            _companyCheckBox.Location = new Point(42, 140);
            _companyCheckBox.Checked = preferences == null || preferences.SyncCompanyDirectory;
            Controls.Add(_companyCheckBox);

            _personalCheckBox.Text = "Persönliche Kontakte synchronisieren";
            _personalCheckBox.AutoSize = true;
            _personalCheckBox.Location = new Point(42, 168);
            _personalCheckBox.Checked = preferences == null || preferences.SyncPersonalContacts;
            Controls.Add(_personalCheckBox);

            var okButton = new Button
            {
                Text = "Übernehmen",
                DialogResult = DialogResult.OK,
                Location = new Point(280, 205),
                Size = new Size(105, 30)
            };
            Controls.Add(okButton);

            var cancelButton = new Button
            {
                Text = "Nicht synchronisieren",
                DialogResult = DialogResult.Cancel,
                Location = new Point(395, 205),
                Size = new Size(105, 30)
            };
            Controls.Add(cancelButton);

            AcceptButton = okButton;
            CancelButton = cancelButton;
            UpdateState();
            UiThemeManager.ApplyToForm(this, null);
        }

        internal void ApplyTo(CardDavSyncPreferences preferences, bool accepted)
        {
            if (preferences == null)
            {
                return;
            }

            preferences.Configured = true;
            if (!accepted)
            {
                preferences.Enabled = false;
                return;
            }

            preferences.Enabled = _enabledCheckBox.Checked;
            preferences.SyncCompanyDirectory = _companyCheckBox.Checked;
            preferences.SyncPersonalContacts = _personalCheckBox.Checked;
        }

        private void UpdateState()
        {
            _companyCheckBox.Enabled = _enabledCheckBox.Checked;
            _personalCheckBox.Enabled = _enabledCheckBox.Checked;
        }
    }
}
