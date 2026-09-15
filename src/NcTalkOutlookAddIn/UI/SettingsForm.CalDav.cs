// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
    internal sealed partial class SettingsForm
    {
        private readonly GroupBox _calDavSyncGroup = new GroupBox();
        private readonly CheckBox _calDavEnabledCheckBox = new CheckBox();
        private readonly CheckedListBox _calDavCalendarList = new CheckedListBox();
        private readonly Button _calDavRefreshButton = new Button();
        private readonly RadioButton _calDavTwoWayRadio = new RadioButton();
        private readonly RadioButton _calDavNextcloudToOutlookRadio = new RadioButton();
        private readonly RadioButton _calDavOutlookToNextcloudRadio = new RadioButton();
        private readonly RadioButton _calDavSeparateFoldersRadio = new RadioButton();
        private readonly RadioButton _calDavDefaultCalendarRadio = new RadioButton();
        private readonly CheckBox _calDavStartupCheckBox = new CheckBox();
        private CalDavSyncPreferences _calDavPreferences;
        private IList<CalDavCalendar> _calDavCalendars = new List<CalDavCalendar>();

        private void InitializeCalDavSettingsSection()
        {
            _calDavPreferences = CalDavSyncPreferences.Load();

            _calDavSyncGroup.Text = "Nextcloud-Kalender";
            _calDavSyncGroup.Location = new Point(18, 515);
            _calDavSyncGroup.Size = new Size(440, 430);
            _calDavSyncGroup.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            _generalTab.Controls.Add(_calDavSyncGroup);

            _calDavEnabledCheckBox.Text = "Kalendersynchronisation aktiv";
            _calDavEnabledCheckBox.AutoSize = true;
            _calDavEnabledCheckBox.Location = new Point(12, 24);
            _calDavEnabledCheckBox.Checked = _calDavPreferences.Enabled;
            _calDavEnabledCheckBox.CheckedChanged += delegate { UpdateCalDavSettingsState(); };
            _calDavSyncGroup.Controls.Add(_calDavEnabledCheckBox);

            _calDavRefreshButton.Text = "Kalender laden / aktualisieren";
            _calDavRefreshButton.Location = new Point(235, 20);
            _calDavRefreshButton.Size = new Size(190, 28);
            _calDavRefreshButton.Click += OnCalDavRefreshClick;
            _calDavSyncGroup.Controls.Add(_calDavRefreshButton);

            _calDavCalendarList.CheckOnClick = true;
            _calDavCalendarList.Location = new Point(30, 58);
            _calDavCalendarList.Size = new Size(395, 115);
            _calDavSyncGroup.Controls.Add(_calDavCalendarList);

            var directionLabel = new Label
            {
                Text = "Synchronisationsrichtung:",
                AutoSize = true,
                Location = new Point(30, 182)
            };
            _calDavSyncGroup.Controls.Add(directionLabel);

            _calDavTwoWayRadio.Text = "Bidirektional (Outlook ↔ Nextcloud)";
            _calDavTwoWayRadio.AutoSize = true;
            _calDavTwoWayRadio.Location = new Point(45, 204);
            _calDavTwoWayRadio.Checked = _calDavPreferences.Direction == CalDavSyncDirection.TwoWay;
            _calDavSyncGroup.Controls.Add(_calDavTwoWayRadio);

            _calDavNextcloudToOutlookRadio.Text = "Nur Nextcloud → Outlook";
            _calDavNextcloudToOutlookRadio.AutoSize = true;
            _calDavNextcloudToOutlookRadio.Location = new Point(45, 228);
            _calDavNextcloudToOutlookRadio.Checked = _calDavPreferences.Direction == CalDavSyncDirection.NextcloudToOutlook;
            _calDavSyncGroup.Controls.Add(_calDavNextcloudToOutlookRadio);

            _calDavOutlookToNextcloudRadio.Text = "Nur Outlook → Nextcloud";
            _calDavOutlookToNextcloudRadio.AutoSize = true;
            _calDavOutlookToNextcloudRadio.Location = new Point(45, 252);
            _calDavOutlookToNextcloudRadio.Checked = _calDavPreferences.Direction == CalDavSyncDirection.OutlookToNextcloud;
            _calDavSyncGroup.Controls.Add(_calDavOutlookToNextcloudRadio);

            var targetLabel = new Label
            {
                Text = "Ziel in Outlook:",
                AutoSize = true,
                Location = new Point(30, 282)
            };
            _calDavSyncGroup.Controls.Add(targetLabel);

            _calDavSeparateFoldersRadio.Text = "Eigener Outlook-Kalender pro Nextcloud-Kalender";
            _calDavSeparateFoldersRadio.AutoSize = true;
            _calDavSeparateFoldersRadio.Location = new Point(45, 304);
            _calDavSeparateFoldersRadio.Checked = !_calDavPreferences.UseDefaultCalendarFolder;
            _calDavSyncGroup.Controls.Add(_calDavSeparateFoldersRadio);

            _calDavDefaultCalendarRadio.Text = "Allgemeiner Outlook-Kalender (Nur dieser Computer)";
            _calDavDefaultCalendarRadio.AutoSize = true;
            _calDavDefaultCalendarRadio.Location = new Point(45, 328);
            _calDavDefaultCalendarRadio.Checked = _calDavPreferences.UseDefaultCalendarFolder;
            _calDavSyncGroup.Controls.Add(_calDavDefaultCalendarRadio);

            _calDavStartupCheckBox.Text = "Beim Outlook-Start sofort synchronisieren";
            _calDavStartupCheckBox.AutoSize = true;
            _calDavStartupCheckBox.Location = new Point(30, 360);
            _calDavStartupCheckBox.Checked = _calDavPreferences.SyncOnStartup;
            _calDavSyncGroup.Controls.Add(_calDavStartupCheckBox);

            _calDavSyncGroup.Controls.Add(new Label
            {
                Text = "Hintergrundaktualisierung: alle 15 Minuten",
                AutoSize = true,
                Location = new Point(30, 389)
            });

            PopulateSavedCalDavCalendarSelection();
            FormClosed += OnCalDavSettingsFormClosed;
            UpdateCalDavSettingsState();
        }

        private void PopulateSavedCalDavCalendarSelection()
        {
            _calDavCalendarList.Items.Clear();
            foreach (string href in _calDavPreferences.SelectedCalendarHrefs)
            {
                _calDavCalendarList.Items.Add(href, true);
            }
        }

        private async void OnCalDavRefreshClick(object sender, EventArgs e)
        {
            if (_isBusy)
            {
                return;
            }

            var configuration = new TalkServiceConfiguration(
                _serverUrlTextBox.Text.Trim(),
                _usernameTextBox.Text.Trim(),
                _appPasswordTextBox.Text ?? string.Empty);
            if (!configuration.IsComplete())
            {
                SetStatus("Nextcloud-Zugangsdaten sind unvollständig.", true);
                return;
            }

            var selectedBeforeRefresh = GetSelectedCalDavHrefs();
            if (selectedBeforeRefresh.Count == 0)
            {
                selectedBeforeRefresh = new List<string>(_calDavPreferences.SelectedCalendarHrefs);
            }

            SetBusy(true);
            UpdateCalDavSettingsState();
            SetStatus("Nextcloud-Kalender werden geladen ...", false);
            try
            {
                _calDavCalendars = await Task.Run(() =>
                    new DavDiscoveryService(configuration).DiscoverCalendars());

                _calDavCalendarList.Items.Clear();
                foreach (CalDavCalendar calendar in _calDavCalendars)
                {
                    string label = string.IsNullOrWhiteSpace(calendar.DisplayName)
                        ? calendar.Href
                        : calendar.DisplayName;
                    if (calendar.ReadOnly)
                    {
                        label += " (nur lesen)";
                    }

                    bool selected = selectedBeforeRefresh.Any(v =>
                        string.Equals(v, calendar.Href, StringComparison.OrdinalIgnoreCase));
                    _calDavCalendarList.Items.Add(label, selected);
                }
                SetStatus("Nextcloud-Kalender geladen: " + _calDavCalendars.Count + ".", false);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "CalDAV calendar discovery from settings failed.", ex);
                SetStatus("Kalender konnten nicht geladen werden: " + ex.Message, true);
            }
            finally
            {
                SetBusy(false);
                UpdateCalDavSettingsState();
            }
        }

        private List<string> GetSelectedCalDavHrefs()
        {
            var result = new List<string>();
            if (_calDavCalendars.Count > 0)
            {
                for (int i = 0; i < _calDavCalendarList.Items.Count && i < _calDavCalendars.Count; i++)
                {
                    if (_calDavCalendarList.GetItemChecked(i))
                    {
                        result.Add(_calDavCalendars[i].Href);
                    }
                }
                return result;
            }

            for (int i = 0; i < _calDavCalendarList.Items.Count; i++)
            {
                if (_calDavCalendarList.GetItemChecked(i))
                {
                    result.Add(Convert.ToString(_calDavCalendarList.Items[i]) ?? string.Empty);
                }
            }
            return result;
        }

        private void UpdateCalDavSettingsState()
        {
            bool enabled = _calDavEnabledCheckBox.Checked;
            _calDavCalendarList.Enabled = enabled;
            _calDavRefreshButton.Enabled = enabled && !_isBusy;
            _calDavTwoWayRadio.Enabled = enabled;
            _calDavNextcloudToOutlookRadio.Enabled = enabled;
            _calDavOutlookToNextcloudRadio.Enabled = enabled;
            _calDavSeparateFoldersRadio.Enabled = enabled;
            _calDavDefaultCalendarRadio.Enabled = enabled;
            _calDavStartupCheckBox.Enabled = enabled;
        }

        private void OnCalDavSettingsFormClosed(object sender, FormClosedEventArgs e)
        {
            if (DialogResult != DialogResult.OK || _calDavPreferences == null)
            {
                return;
            }

            try
            {
                _calDavPreferences.Configured = true;
                _calDavPreferences.Enabled = _calDavEnabledCheckBox.Checked;
                _calDavPreferences.Direction = _calDavNextcloudToOutlookRadio.Checked
                    ? CalDavSyncDirection.NextcloudToOutlook
                    : (_calDavOutlookToNextcloudRadio.Checked
                        ? CalDavSyncDirection.OutlookToNextcloud
                        : CalDavSyncDirection.TwoWay);
                _calDavPreferences.UseDefaultCalendarFolder = _calDavDefaultCalendarRadio.Checked;
                _calDavPreferences.SyncOnStartup = _calDavStartupCheckBox.Checked;
                _calDavPreferences.IntervalMinutes = 15;
                _calDavPreferences.SetSelectedCalendars(GetSelectedCalDavHrefs());
                _calDavPreferences.Save();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to save CalDAV settings from settings form.", ex);
            }
        }
    }
}
