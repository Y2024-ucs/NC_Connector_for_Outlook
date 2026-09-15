// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
    internal sealed class CalDavFirstRunForm : ScaledForm
    {
        private readonly IList<CalDavCalendar> _calendars;
        private readonly CheckBox _enabledCheckBox = new CheckBox();
        private readonly CheckedListBox _calendarList = new CheckedListBox();
        private readonly RadioButton _twoWayRadio = new RadioButton();
        private readonly RadioButton _nextcloudToOutlookRadio = new RadioButton();
        private readonly RadioButton _outlookToNextcloudRadio = new RadioButton();
        private readonly RadioButton _separateFoldersRadio = new RadioButton();
        private readonly RadioButton _defaultCalendarRadio = new RadioButton();
        private readonly CheckBox _startupCheckBox = new CheckBox();

        internal CalDavFirstRunForm(
            IList<CalDavCalendar> calendars,
            CalDavSyncPreferences preferences)
        {
            _calendars = calendars ?? new List<CalDavCalendar>();

            Text = "Nextcloud-Kalender synchronisieren";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(650, 590);
            Icon = BrandingAssets.GetAppIcon(32);

            Controls.Add(new Label
            {
                Text = "Welche Nextcloud-Kalender sollen mit Outlook synchronisiert werden?",
                AutoSize = false,
                Font = new Font(Font, FontStyle.Bold),
                Location = new Point(20, 18),
                Size = new Size(610, 32)
            });

            Controls.Add(new Label
            {
                Text = "Standard ist bidirektional. Schreibgeschützte Kalender werden auf der Nextcloud-Seite niemals verändert.",
                AutoSize = false,
                Location = new Point(20, 50),
                Size = new Size(610, 38)
            });

            _enabledCheckBox.Text = "Kalendersynchronisation aktivieren";
            _enabledCheckBox.AutoSize = true;
            _enabledCheckBox.Location = new Point(20, 92);
            _enabledCheckBox.Checked = preferences == null || preferences.Enabled;
            _enabledCheckBox.CheckedChanged += delegate { UpdateState(); };
            Controls.Add(_enabledCheckBox);

            Controls.Add(new Label
            {
                Text = "Kalender:",
                AutoSize = true,
                Font = new Font(Font, FontStyle.Bold),
                Location = new Point(20, 124)
            });

            _calendarList.CheckOnClick = true;
            _calendarList.Location = new Point(20, 148);
            _calendarList.Size = new Size(610, 150);
            for (int i = 0; i < _calendars.Count; i++)
            {
                CalDavCalendar calendar = _calendars[i];
                string label = string.IsNullOrWhiteSpace(calendar.DisplayName)
                    ? calendar.Href
                    : calendar.DisplayName;
                if (calendar.ReadOnly)
                {
                    label += " (nur lesen)";
                }
                bool selected = preferences == null
                    || preferences.SelectedCalendarHrefs.Count == 0
                    || preferences.IsCalendarSelected(calendar.Href);
                _calendarList.Items.Add(label, selected);
            }
            Controls.Add(_calendarList);

            Controls.Add(new Label
            {
                Text = "Synchronisationsrichtung:",
                AutoSize = true,
                Font = new Font(Font, FontStyle.Bold),
                Location = new Point(20, 314)
            });

            _twoWayRadio.Text = "Bidirektional (Outlook ↔ Nextcloud)";
            _twoWayRadio.AutoSize = true;
            _twoWayRadio.Location = new Point(42, 340);
            _twoWayRadio.Checked = preferences == null || preferences.Direction == CalDavSyncDirection.TwoWay;
            Controls.Add(_twoWayRadio);

            _nextcloudToOutlookRadio.Text = "Nur Nextcloud → Outlook";
            _nextcloudToOutlookRadio.AutoSize = true;
            _nextcloudToOutlookRadio.Location = new Point(42, 366);
            _nextcloudToOutlookRadio.Checked = preferences != null && preferences.Direction == CalDavSyncDirection.NextcloudToOutlook;
            Controls.Add(_nextcloudToOutlookRadio);

            _outlookToNextcloudRadio.Text = "Nur Outlook → Nextcloud";
            _outlookToNextcloudRadio.AutoSize = true;
            _outlookToNextcloudRadio.Location = new Point(42, 392);
            _outlookToNextcloudRadio.Checked = preferences != null && preferences.Direction == CalDavSyncDirection.OutlookToNextcloud;
            Controls.Add(_outlookToNextcloudRadio);

            Controls.Add(new Label
            {
                Text = "Ziel in Outlook:",
                AutoSize = true,
                Font = new Font(Font, FontStyle.Bold),
                Location = new Point(20, 425)
            });

            _separateFoldersRadio.Text = "Eigener Outlook-Kalender pro Nextcloud-Kalender";
            _separateFoldersRadio.AutoSize = true;
            _separateFoldersRadio.Location = new Point(42, 451);
            _separateFoldersRadio.Checked = preferences == null || !preferences.UseDefaultCalendarFolder;
            Controls.Add(_separateFoldersRadio);

            _defaultCalendarRadio.Text = "Allgemeiner Outlook-Kalender (Kalender - Nur dieser Computer)";
            _defaultCalendarRadio.AutoSize = true;
            _defaultCalendarRadio.Location = new Point(42, 477);
            _defaultCalendarRadio.Checked = preferences != null && preferences.UseDefaultCalendarFolder;
            Controls.Add(_defaultCalendarRadio);

            _startupCheckBox.Text = "Beim Outlook-Start sofort synchronisieren (danach alle 15 Minuten)";
            _startupCheckBox.AutoSize = true;
            _startupCheckBox.Location = new Point(20, 514);
            _startupCheckBox.Checked = preferences == null || preferences.SyncOnStartup;
            Controls.Add(_startupCheckBox);

            var okButton = new Button
            {
                Text = "Übernehmen",
                DialogResult = DialogResult.OK,
                Location = new Point(405, 548),
                Size = new Size(105, 30)
            };
            Controls.Add(okButton);

            var cancelButton = new Button
            {
                Text = "Nicht synchronisieren",
                DialogResult = DialogResult.Cancel,
                Location = new Point(520, 548),
                Size = new Size(110, 30)
            };
            Controls.Add(cancelButton);

            AcceptButton = okButton;
            CancelButton = cancelButton;
            UpdateState();
            UiThemeManager.ApplyToForm(this, null);
        }

        internal void ApplyTo(CalDavSyncPreferences preferences, bool accepted)
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
            preferences.Direction = _nextcloudToOutlookRadio.Checked
                ? CalDavSyncDirection.NextcloudToOutlook
                : (_outlookToNextcloudRadio.Checked
                    ? CalDavSyncDirection.OutlookToNextcloud
                    : CalDavSyncDirection.TwoWay);
            preferences.UseDefaultCalendarFolder = _defaultCalendarRadio.Checked;
            preferences.SyncOnStartup = _startupCheckBox.Checked;
            preferences.IntervalMinutes = 15;

            var selectedHrefs = new List<string>();
            for (int i = 0; i < _calendarList.Items.Count && i < _calendars.Count; i++)
            {
                if (_calendarList.GetItemChecked(i))
                {
                    selectedHrefs.Add(_calendars[i].Href);
                }
            }
            preferences.SetSelectedCalendars(selectedHrefs);
        }

        private void UpdateState()
        {
            bool enabled = _enabledCheckBox.Checked;
            _calendarList.Enabled = enabled;
            _twoWayRadio.Enabled = enabled;
            _nextcloudToOutlookRadio.Enabled = enabled;
            _outlookToNextcloudRadio.Enabled = enabled;
            _separateFoldersRadio.Enabled = enabled;
            _defaultCalendarRadio.Enabled = enabled;
            _startupCheckBox.Enabled = enabled;
        }
    }
}
