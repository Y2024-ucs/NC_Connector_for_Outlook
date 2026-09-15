// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
    internal sealed partial class SettingsForm
    {
        private readonly TabPage _cardDavTab = new TabPage("Kontakte");
        private readonly GroupBox _cardDavSyncGroup = new GroupBox();
        private readonly CheckBox _cardDavEnabledCheckBox = new CheckBox();
        private readonly CheckBox _cardDavCompanyCheckBox = new CheckBox();
        private readonly CheckBox _cardDavPersonalCheckBox = new CheckBox();
        private readonly RadioButton _cardDavSeparateFoldersRadio = new RadioButton();
        private readonly RadioButton _cardDavDefaultContactsRadio = new RadioButton();
        private readonly Button _cardDavSyncNowButton = new Button();
        private CardDavSyncPreferences _cardDavPreferences;

        private void InitializeCardDavSettingsSection()
        {
            _cardDavPreferences = CardDavSyncPreferences.Load();

            _cardDavTab.AutoScroll = true;
            if (!_tabControl.TabPages.Contains(_cardDavTab))
            {
                _tabControl.TabPages.Insert(1, _cardDavTab);
            }

            _cardDavSyncGroup.Text = "Nextcloud-Kontakte";
            _cardDavSyncGroup.Location = new Point(18, 20);
            _cardDavSyncGroup.Size = new Size(700, 230);
            _cardDavSyncGroup.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            _cardDavTab.Controls.Add(_cardDavSyncGroup);

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
            _cardDavCompanyCheckBox.CheckedChanged += delegate { UpdateCardDavSettingsState(); };
            _cardDavSyncGroup.Controls.Add(_cardDavCompanyCheckBox);

            _cardDavPersonalCheckBox.Text = "Persönliche Kontakte synchronisieren";
            _cardDavPersonalCheckBox.AutoSize = true;
            _cardDavPersonalCheckBox.Location = new Point(30, 82);
            _cardDavPersonalCheckBox.Checked = _cardDavPreferences.SyncPersonalContacts;
            _cardDavPersonalCheckBox.CheckedChanged += delegate { UpdateCardDavSettingsState(); };
            _cardDavSyncGroup.Controls.Add(_cardDavPersonalCheckBox);

            var targetLabel = new Label
            {
                Text = "Ziel in Outlook:",
                AutoSize = true,
                Location = new Point(30, 112)
            };
            _cardDavSyncGroup.Controls.Add(targetLabel);

            _cardDavSeparateFoldersRadio.Text = "Eigene Nextcloud-Kontaktordner";
            _cardDavSeparateFoldersRadio.AutoSize = true;
            _cardDavSeparateFoldersRadio.Location = new Point(45, 136);
            _cardDavSeparateFoldersRadio.Checked = !_cardDavPreferences.UseDefaultContactsFolder;
            _cardDavSyncGroup.Controls.Add(_cardDavSeparateFoldersRadio);

            _cardDavDefaultContactsRadio.Text = "Allgemeine Outlook-Kontakte (Nur dieser Computer)";
            _cardDavDefaultContactsRadio.AutoSize = true;
            _cardDavDefaultContactsRadio.Location = new Point(45, 160);
            _cardDavDefaultContactsRadio.Checked = _cardDavPreferences.UseDefaultContactsFolder;
            _cardDavSyncGroup.Controls.Add(_cardDavDefaultContactsRadio);

            _cardDavSyncNowButton.Text = "Jetzt synchronisieren";
            _cardDavSyncNowButton.Location = new Point(30, 187);
            _cardDavSyncNowButton.Size = new Size(170, 28);
            _cardDavSyncNowButton.Click += OnCardDavSyncNowClick;
            _cardDavSyncGroup.Controls.Add(_cardDavSyncNowButton);

            var hint = new Label
            {
                Text = "Nur lesend: Nextcloud → Outlook · nur geänderte Kontakte · Hintergrundaktualisierung alle 15 Minuten",
                AutoSize = true,
                Location = new Point(215, 194)
            };
            _cardDavSyncGroup.Controls.Add(hint);

            FormClosed += OnCardDavSettingsFormClosed;
            UpdateCardDavSettingsState();
        }

        private void UpdateCardDavSettingsState()
        {
            bool enabled = _cardDavEnabledCheckBox.Checked;
            bool hasSelection = _cardDavCompanyCheckBox.Checked || _cardDavPersonalCheckBox.Checked;
            _cardDavCompanyCheckBox.Enabled = enabled;
            _cardDavPersonalCheckBox.Enabled = enabled;
            _cardDavSeparateFoldersRadio.Enabled = enabled;
            _cardDavDefaultContactsRadio.Enabled = enabled;
            _cardDavSyncNowButton.Enabled = enabled && hasSelection && !_isBusy;
        }

        private async void OnCardDavSyncNowClick(object sender, EventArgs e)
        {
            if (_isBusy || _outlookApplication == null)
            {
                return;
            }

            string baseUrl = _serverUrlTextBox.Text.Trim();
            string user = _usernameTextBox.Text.Trim();
            string appPassword = _appPasswordTextBox.Text ?? string.Empty;
            var configuration = new TalkServiceConfiguration(baseUrl, user, appPassword);
            if (!configuration.IsComplete())
            {
                SetStatus("Nextcloud-Zugangsdaten sind unvollständig.", true);
                return;
            }

            bool syncCompany = _cardDavCompanyCheckBox.Checked;
            bool syncPersonal = _cardDavPersonalCheckBox.Checked;
            if (!syncCompany && !syncPersonal)
            {
                SetStatus("Es ist kein Kontaktbereich zur Synchronisation ausgewählt.", true);
                return;
            }

            SetBusy(true);
            UpdateCardDavSettingsState();
            SetStatus("Nextcloud-Kontakte werden auf Änderungen geprüft ...", false);
            try
            {
                Dictionary<string, string> knownEtags = CardDavReadOnlySync.LoadKnownEtagsFromOutlook(_outlookApplication);
                var sync = new CardDavReadOnlySync(configuration);
                List<CardDavContactRecord> contacts = await Task.Run(() =>
                {
                    IList<CardDavAddressBook> addressBooks =
                        new DavDiscoveryService(configuration).DiscoverAddressBooks();
                    var result = new List<CardDavContactRecord>();
                    foreach (CardDavAddressBook addressBook in addressBooks)
                    {
                        bool isSystem = addressBook != null
                            && !string.IsNullOrWhiteSpace(addressBook.Href)
                            && addressBook.Href.IndexOf(
                                "z-server-generated--system",
                                StringComparison.OrdinalIgnoreCase) >= 0;
                        if ((isSystem && !syncCompany) || (!isSystem && !syncPersonal))
                        {
                            continue;
                        }
                        result.AddRange(sync.DownloadContacts(addressBook, knownEtags));
                    }
                    return result;
                });

                int count = _cardDavDefaultContactsRadio.Checked
                    ? CardDavDefaultContactsImporter.Import(_outlookApplication, contacts)
                    : sync.ImportIntoOutlook(_outlookApplication, contacts);
                SetStatus("Nextcloud-Kontakte synchronisiert: " + count + " geändert/neu.", false);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Manual CardDAV sync failed.", ex);
                SetStatus("Kontaktsynchronisation fehlgeschlagen: " + ex.Message, true);
            }
            finally
            {
                SetBusy(false);
                UpdateCardDavSettingsState();
            }
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
                _cardDavPreferences.UseDefaultContactsFolder = _cardDavDefaultContactsRadio.Checked;
                _cardDavPreferences.Save();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to save CardDAV settings from settings form.", ex);
            }
        }
    }
}
