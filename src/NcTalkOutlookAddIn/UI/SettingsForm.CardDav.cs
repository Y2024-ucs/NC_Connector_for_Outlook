// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
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
        private readonly CheckBox _cardDavPersonalCheckBox = new CheckBox();
        private readonly ComboBox _cardDavAddressBookCombo = new ComboBox();
        private readonly Button _cardDavAddressBookReloadButton = new Button();
        private readonly RadioButton _cardDavSeparateFoldersRadio = new RadioButton();
        private readonly RadioButton _cardDavDefaultContactsRadio = new RadioButton();
        private readonly CheckBox _cardDavCustomersCheckBox = new CheckBox();
        private readonly Button _cardDavSyncNowButton = new Button();
        private readonly Button _cardDavResetButton = new Button();
        private CardDavSyncPreferences _cardDavPreferences;
        private bool _cardDavAddressBooksLoading;

        private sealed class AddressBookChoice
        {
            internal string Href { get; set; }
            internal string Name { get; set; }

            public override string ToString()
            {
                return Name;
            }
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            if (_cardDavPreferences == null)
            {
                InitializeCardDavSettingsSection();
            }

            // Keep the dedicated contacts tab visible even if another settings initialization step
            // has rebuilt or reordered the TabPages collection after InitializeGeneralTab().
            if (!_tabControl.TabPages.Contains(_cardDavTab))
            {
                int insertIndex = Math.Min(1, _tabControl.TabPages.Count);
                _tabControl.TabPages.Insert(insertIndex, _cardDavTab);
            }

            var tabNames = new List<string>();
            foreach (TabPage page in _tabControl.TabPages)
            {
                tabNames.Add(page != null ? page.Text : string.Empty);
            }
            DiagnosticsLogger.Log(
                LogCategories.Core,
                "Settings tabs initialized (count=" + _tabControl.TabPages.Count
                + ", tabs=" + string.Join(" | ", tabNames.ToArray()) + ").");
        }

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
            _cardDavSyncGroup.Size = new Size(700, 292);
            _cardDavSyncGroup.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            _cardDavTab.Controls.Add(_cardDavSyncGroup);

            _cardDavEnabledCheckBox.Text = "Kontaktsynchronisation aktiv";
            _cardDavEnabledCheckBox.AutoSize = true;
            _cardDavEnabledCheckBox.Location = new Point(12, 25);
            _cardDavEnabledCheckBox.Checked = _cardDavPreferences.Enabled;
            _cardDavEnabledCheckBox.CheckedChanged += delegate { UpdateCardDavSettingsState(); };
            _cardDavSyncGroup.Controls.Add(_cardDavEnabledCheckBox);

            _cardDavPersonalCheckBox.Text = "Persönliche Kontakte synchronisieren";
            _cardDavPersonalCheckBox.AutoSize = true;
            _cardDavPersonalCheckBox.Location = new Point(30, 55);
            _cardDavPersonalCheckBox.Checked = _cardDavPreferences.SyncPersonalContacts;
            _cardDavPersonalCheckBox.CheckedChanged += delegate { UpdateCardDavSettingsState(); };
            _cardDavSyncGroup.Controls.Add(_cardDavPersonalCheckBox);

            var addressBookLabel = new Label
            {
                Text = "Adressbuch:",
                AutoSize = true,
                Location = new Point(48, 86)
            };
            _cardDavSyncGroup.Controls.Add(addressBookLabel);

            _cardDavAddressBookCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            _cardDavAddressBookCombo.Location = new Point(135, 82);
            _cardDavAddressBookCombo.Size = new Size(300, 24);
            _cardDavSyncGroup.Controls.Add(_cardDavAddressBookCombo);
            FillCardDavAddressBookChoices(null);

            _cardDavAddressBookReloadButton.Text = "Liste laden";
            _cardDavAddressBookReloadButton.Location = new Point(445, 81);
            _cardDavAddressBookReloadButton.Size = new Size(110, 26);
            _cardDavAddressBookReloadButton.Click += delegate { StartLoadingCardDavAddressBooks(true); };
            _cardDavSyncGroup.Controls.Add(_cardDavAddressBookReloadButton);

            var targetLabel = new Label
            {
                Text = "Ziel in Outlook:",
                AutoSize = true,
                Location = new Point(30, 116)
            };
            _cardDavSyncGroup.Controls.Add(targetLabel);

            _cardDavSeparateFoldersRadio.Text = "Eigene Nextcloud-Kontaktordner";
            _cardDavSeparateFoldersRadio.AutoSize = true;
            _cardDavSeparateFoldersRadio.Location = new Point(45, 138);
            _cardDavSeparateFoldersRadio.Checked = !_cardDavPreferences.UseDefaultContactsFolder;
            _cardDavSyncGroup.Controls.Add(_cardDavSeparateFoldersRadio);

            _cardDavDefaultContactsRadio.Text = "Allgemeine Outlook-Kontakte (Nur dieser Computer)";
            _cardDavDefaultContactsRadio.AutoSize = true;
            _cardDavDefaultContactsRadio.Location = new Point(45, 162);
            _cardDavDefaultContactsRadio.Checked = _cardDavPreferences.UseDefaultContactsFolder;
            _cardDavSyncGroup.Controls.Add(_cardDavDefaultContactsRadio);

            _cardDavCustomersCheckBox.Text = "Outlook-Ordner \"" + CardDavCustomerSync.FolderName
                + "\" beidseitig mit Nextcloud-Adressbuch \"" + CardDavCustomerSync.AddressBookDisplayName + "\" abgleichen";
            _cardDavCustomersCheckBox.AutoSize = true;
            _cardDavCustomersCheckBox.Location = new Point(30, 191);
            _cardDavCustomersCheckBox.Checked = _cardDavPreferences.SyncCustomers;
            _cardDavCustomersCheckBox.CheckedChanged += delegate { UpdateCardDavSettingsState(); };
            _cardDavSyncGroup.Controls.Add(_cardDavCustomersCheckBox);

            _cardDavSyncNowButton.Text = "Jetzt synchronisieren";
            _cardDavSyncNowButton.Location = new Point(30, 221);
            _cardDavSyncNowButton.Size = new Size(170, 28);
            _cardDavSyncNowButton.Click += OnCardDavSyncNowClick;
            _cardDavSyncGroup.Controls.Add(_cardDavSyncNowButton);

            _cardDavResetButton.Text = "Kontakte-Sync zurücksetzen und neu aufbauen";
            _cardDavResetButton.Location = new Point(210, 221);
            _cardDavResetButton.Size = new Size(290, 28);
            _cardDavResetButton.Click += OnCardDavResetClick;
            _cardDavSyncGroup.Controls.Add(_cardDavResetButton);

            var hint = new Label
            {
                Text = "Kontakte nur lesend, Kunden beidseitig · nur Änderungen · automatisch alle 15 Minuten",
                AutoSize = true,
                Location = new Point(30, 260)
            };
            _cardDavSyncGroup.Controls.Add(hint);

            FormClosed += OnCardDavSettingsFormClosed;
            UpdateCardDavSettingsState();
            StartLoadingCardDavAddressBooks(false);
        }

        /// <summary>
        /// Fills the address book list: "all" first, then the server's personal books, or only the saved
        /// selection while the server list is unavailable. Keeps the current selection when possible.
        /// </summary>
        private void FillCardDavAddressBookChoices(IList<CardDavAddressBook> serverBooks)
        {
            AddressBookChoice current = _cardDavAddressBookCombo.SelectedItem as AddressBookChoice;
            string selectedHref = current != null ? current.Href : _cardDavPreferences.PersonalAddressBookHref;

            var choices = new List<AddressBookChoice>
            {
                new AddressBookChoice { Href = CardDavSyncPreferences.AllPersonalAddressBooks, Name = "Alle persönlichen Adressbücher" }
            };
            if (serverBooks != null)
            {
                foreach (CardDavAddressBook book in serverBooks)
                {
                    if (!CardDavSyncPreferences.IsPersonalAddressBook(book))
                    {
                        continue;
                    }
                    choices.Add(new AddressBookChoice
                    {
                        Href = book.Href,
                        Name = string.IsNullOrWhiteSpace(book.DisplayName) ? book.Href : book.DisplayName
                    });
                }
            }
            if (string.IsNullOrWhiteSpace(selectedHref))
            {
                // Nothing chosen yet: the IBP default book, or "all" when the account has none.
                CardDavAddressBook defaultBook = serverBooks == null
                    ? null
                    : serverBooks.FirstOrDefault(b => CardDavSyncPreferences.IsPersonalAddressBook(b)
                        && CardDavSyncPreferences.IsDefaultAddressBook(b));
                if (defaultBook != null)
                {
                    selectedHref = defaultBook.Href;
                }
                else if (serverBooks != null)
                {
                    selectedHref = CardDavSyncPreferences.AllPersonalAddressBooks;
                }
                else
                {
                    choices.Add(new AddressBookChoice
                    {
                        Href = string.Empty,
                        Name = CardDavSyncPreferences.DefaultAddressBookDisplayName + " (Standard)"
                    });
                }
            }
            if (!string.IsNullOrWhiteSpace(selectedHref)
                && !choices.Exists(c => CardDavSyncPreferences.SameCollection(c.Href, selectedHref)))
            {
                string name = string.IsNullOrWhiteSpace(_cardDavPreferences.PersonalAddressBookName)
                    ? selectedHref
                    : _cardDavPreferences.PersonalAddressBookName;
                choices.Add(new AddressBookChoice
                {
                    Href = selectedHref,
                    Name = serverBooks != null ? name + " (nicht mehr vorhanden)" : name
                });
            }

            _cardDavAddressBookCombo.BeginUpdate();
            _cardDavAddressBookCombo.Items.Clear();
            foreach (AddressBookChoice choice in choices)
            {
                _cardDavAddressBookCombo.Items.Add(choice);
            }
            _cardDavAddressBookCombo.SelectedItem = choices.Find(c => CardDavSyncPreferences.SameCollection(c.Href, selectedHref))
                ?? choices[0];
            _cardDavAddressBookCombo.EndUpdate();
        }

        private void StartLoadingCardDavAddressBooks(bool userInitiated)
        {
            LoadCardDavAddressBooksAsync(userInitiated).ContinueWith(
                task => DiagnosticsLogger.LogException(LogCategories.Core, "Loading CardDAV address books failed.", task.Exception),
                CancellationToken.None,
                TaskContinuationOptions.OnlyOnFaulted,
                TaskScheduler.Default);
        }

        private async Task LoadCardDavAddressBooksAsync(bool userInitiated)
        {
            if (_cardDavAddressBooksLoading)
            {
                return;
            }
            var configuration = new TalkServiceConfiguration(
                _serverUrlTextBox.Text.Trim(),
                _usernameTextBox.Text.Trim(),
                _appPasswordTextBox.Text ?? string.Empty);
            if (!configuration.IsComplete())
            {
                if (userInitiated) SetStatus("Nextcloud-Zugangsdaten sind unvollständig.", true);
                return;
            }

            _cardDavAddressBooksLoading = true;
            UpdateCardDavSettingsState();
            try
            {
                IList<CardDavAddressBook> books = await Task.Run(
                    () => new DavDiscoveryService(configuration).DiscoverAddressBooks());
                if (IsDisposed) return;
                FillCardDavAddressBookChoices(books);
                if (userInitiated) SetStatus("Adressbücher geladen.", false);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Loading CardDAV address books for settings failed.", ex);
                if (userInitiated && !IsDisposed) SetStatus("Adressbücher konnten nicht geladen werden: " + ex.Message, true);
            }
            finally
            {
                _cardDavAddressBooksLoading = false;
                if (!IsDisposed) UpdateCardDavSettingsState();
            }
        }

        /// <summary>
        /// Current, unsaved choices of this tab, used by the manual sync buttons.
        /// </summary>
        private CardDavSyncPreferences BuildCardDavPreferencesFromForm()
        {
            AddressBookChoice choice = _cardDavAddressBookCombo.SelectedItem as AddressBookChoice;
            return new CardDavSyncPreferences
            {
                Configured = true,
                Enabled = _cardDavEnabledCheckBox.Checked,
                SyncPersonalContacts = _cardDavPersonalCheckBox.Checked,
                PersonalAddressBookHref = choice != null ? choice.Href : string.Empty,
                PersonalAddressBookName = choice != null && !string.IsNullOrWhiteSpace(choice.Href)
                    && choice.Href != CardDavSyncPreferences.AllPersonalAddressBooks
                    ? choice.Name.Replace(" (nicht mehr vorhanden)", string.Empty)
                    : string.Empty,
                UseDefaultContactsFolder = _cardDavDefaultContactsRadio.Checked,
                SyncCustomers = _cardDavCustomersCheckBox.Checked
            };
        }

        private void UpdateCardDavSettingsState()
        {
            bool enabled = _cardDavEnabledCheckBox.Checked;
            bool hasSelection = _cardDavPersonalCheckBox.Checked
                || _cardDavCustomersCheckBox.Checked;
            _cardDavCustomersCheckBox.Enabled = enabled;
            _cardDavAddressBookCombo.Enabled = enabled && _cardDavPersonalCheckBox.Checked;
            _cardDavAddressBookReloadButton.Enabled = enabled && !_cardDavAddressBooksLoading && !_isBusy;
            _cardDavPersonalCheckBox.Enabled = enabled;
            _cardDavSeparateFoldersRadio.Enabled = enabled;
            _cardDavDefaultContactsRadio.Enabled = enabled;
            _cardDavSyncNowButton.Enabled = enabled && hasSelection && !_isBusy;
            _cardDavResetButton.Enabled = enabled && _cardDavPersonalCheckBox.Checked && !_isBusy;
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

            CardDavSyncPreferences preferences = BuildCardDavPreferencesFromForm();
            bool syncPersonal = preferences.SyncPersonalContacts;
            bool syncCustomers = preferences.SyncCustomers;
            if (!syncPersonal && !syncCustomers)
            {
                SetStatus("Es ist kein Kontaktbereich zur Synchronisation ausgewählt.", true);
                return;
            }

            if (!CardDavBackgroundSyncManager.TryBeginSync())
            {
                SetStatus("Die Kontaktsynchronisation läuft bereits.", false);
                return;
            }
            SetBusy(true);
            UpdateCardDavSettingsState();
            SetStatus("Nextcloud-Kontakte werden auf Änderungen geprüft ...", false);
            try
            {
                string status = string.Empty;
                if (syncPersonal)
                {
                    status = await RunCardDavReadOnlySyncAsync(configuration, preferences);
                }
                if (syncCustomers)
                {
                    SetStatus("Kunden werden abgeglichen ...", false);
                    CardDavCustomerSyncResult customers = await new CardDavCustomerSync(
                        configuration,
                        _outlookApplication,
                        RunOnSettingsUiThreadAsync,
                        () => Task.Delay(20)).RunAsync();
                    status = (status + " " + customers.Summary()).Trim();
                }
                SetStatus(status, false);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Manual CardDAV sync failed.", ex);
                SetStatus("Kontaktsynchronisation fehlgeschlagen: " + ex.Message, true);
            }
            finally
            {
                CardDavBackgroundSyncManager.EndSync();
                SetBusy(false);
                UpdateCardDavSettingsState();
            }
        }

        private async void OnCardDavResetClick(object sender, EventArgs e)
        {
            if (_isBusy || _outlookApplication == null)
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

            CardDavSyncPreferences preferences = BuildCardDavPreferencesFromForm();
            if (!preferences.SyncPersonalContacts)
            {
                SetStatus("Es ist kein Kontaktbereich zur Synchronisation ausgewählt.", true);
                return;
            }

            DialogResult answer = MessageBox.Show(
                this,
                "Alle aus Nextcloud importierten Kontakte, die Gruppenordner und die Gruppen-Verteilerlisten werden "
                + "aus Outlook entfernt (sie landen in \"Gelöschte Elemente\") und danach vollständig neu geladen."
                + Environment.NewLine + Environment.NewLine
                + "Eigene Outlook-Kontakte und der Ordner \"" + CardDavCustomerSync.FolderName + "\" bleiben unverändert."
                + Environment.NewLine + Environment.NewLine
                + "Fortfahren?",
                "Kontakte-Sync zurücksetzen",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2);
            if (answer != DialogResult.Yes)
            {
                return;
            }

            if (!CardDavBackgroundSyncManager.TryBeginSync())
            {
                SetStatus("Die Kontaktsynchronisation läuft bereits.", false);
                return;
            }
            SetBusy(true);
            UpdateCardDavSettingsState();
            try
            {
                SetStatus("Kontakte-Sync wird zurückgesetzt ...", false);
                CardDavContactsResetResult reset = await CardDavContactsReset.RunAsync(
                    _outlookApplication,
                    RunOnSettingsUiThreadAsync,
                    () => Task.Delay(20));
                SetStatus("Kontakte werden neu geladen ...", false);
                string status = await RunCardDavReadOnlySyncAsync(configuration, preferences, true);
                SetStatus("Zurückgesetzt: " + reset.Contacts + " Kontakte, " + reset.DistributionLists
                    + " Verteilerlisten und " + reset.GroupFolders + " Gruppenordner entfernt. " + status, false);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "CardDAV contact sync reset failed.", ex);
                SetStatus("Zurücksetzen fehlgeschlagen: " + ex.Message, true);
            }
            finally
            {
                CardDavBackgroundSyncManager.EndSync();
                SetBusy(false);
                UpdateCardDavSettingsState();
            }
        }

        private async Task<string> RunCardDavReadOnlySyncAsync(
            TalkServiceConfiguration configuration,
            CardDavSyncPreferences preferences,
            bool includeGroups = false)
        {
            bool useDefaultContactsFolder = preferences.UseDefaultContactsFolder;
            Dictionary<string, string> knownEtags = CardDavReadOnlySync.LoadKnownEtagsFromOutlook(
                _outlookApplication,
                useDefaultContactsFolder);
            var sync = new CardDavReadOnlySync(configuration);
            IList<CardDavAddressBook> syncedAddressBooks = null;
            List<CardDavContactRecord> contacts = await Task.Run(() =>
            {
                IList<CardDavAddressBook> addressBooks = CardDavCustomerSync.WithoutCustomerAddressBook(
                    new DavDiscoveryService(configuration).DiscoverAddressBooks());
                syncedAddressBooks = addressBooks;
                var result = new List<CardDavContactRecord>();
                foreach (CardDavAddressBook addressBook in addressBooks)
                {
                    if (!preferences.IncludesAddressBook(addressBook, addressBooks))
                    {
                        continue;
                    }
                    result.AddRange(sync.DownloadContacts(addressBook, knownEtags));
                }
                return result;
            });

            int count = sync.ImportIntoOutlook(_outlookApplication, contacts,
                useDefaultContactsFolder);
            if (!includeGroups)
            {
                return "Nextcloud-Kontakte synchronisiert: " + count + " geändert/neu, " + sync.DeletedCount + " gelöscht.";
            }

            int groups = CardDavContactGroupSync.Reconcile(configuration, _outlookApplication, syncedAddressBooks, preferences);
            int groupFolders = CardDavContactFolderSync.Reconcile(configuration, _outlookApplication, syncedAddressBooks, preferences);
            return "Neu geladen: " + count + " Kontakte, " + groups + " Gruppen und " + groupFolders + " Gruppenordner.";
        }

        /// <summary>
        /// The customer sync resumes on pool threads; its Outlook slices must return to this form's thread.
        /// </summary>
        private Task RunOnSettingsUiThreadAsync(Action callback)
        {
            var completion = new TaskCompletionSource<bool>();
            Action run = () =>
            {
                try
                {
                    callback();
                    completion.TrySetResult(true);
                }
                catch (Exception ex)
                {
                    completion.TrySetException(ex);
                }
            };

            if (!InvokeRequired)
            {
                run();
                return completion.Task;
            }
            try
            {
                BeginInvoke(run);
            }
            catch (Exception ex)
            {
                // The form was closed while the sync was still running.
                completion.TrySetException(ex);
            }
            return completion.Task;
        }

        private void OnCardDavSettingsFormClosed(object sender, FormClosedEventArgs e)
        {
            if (DialogResult != DialogResult.OK || _cardDavPreferences == null)
            {
                return;
            }

            try
            {
                BuildCardDavPreferencesFromForm().Save();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to save CardDAV settings from settings form.", ex);
            }
        }
    }
}
