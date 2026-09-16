# CardDAV: Ribbon and deletion reconciliation

The Nextcloud Explorer ribbon includes **Kontakte synchronisieren** only while that window displays a contacts folder. Folder switches invalidate visibility. The action uses the configured company/personal selection and reports updated and deleted counts. Startup, background and manual runs share a sync guard.

Deletion uses a complete `207 DAV:multistatus` addressbook-query snapshot, including unchanged contacts. Missing/failed properties, DAV errors (including result limits), invalid/foreign/duplicate hrefs and incomplete payload responses abort the run before reconciliation. A successful empty address book removes its last managed contacts. Undiscovered or unselected address books never authorize deletions.

Only contacts with an `NC-CardDAV-HREF` in an explicitly queried collection on the same origin are candidates. Reconciliation scans the default contacts folder and its immediate contact subfolders, matching the existing import/ETag scan scope. It uses Outlook `ContactItem.Delete()` on the UI thread and reverse item traversal. Ordinary contacts without a marker remain untouched. The sync remains one-way: Nextcloud to Outlook.

## Automated checks

Run `tools/ci/Invoke-OutlookCardDavTests.ps1` on Windows. This compiles the production snapshot parser and tests deletion boundaries, empty books, unchanged contacts, URI normalization and failed/incomplete responses. The workflow also builds the add-in and checks ribbon callbacks and event hook symmetry.

## Outlook acceptance checks

1. Open Mail, Calendar, Contacts and an empty contacts subfolder; the button appears only for contacts. Repeat with two Explorer windows and an appointment Inspector open.
2. Sync from the ribbon; check the completion message and `deleted=N` log. Repeat from Settings, startup and the 15-minute background timer.
3. Import two contacts, leave one unchanged and delete one in Nextcloud. The unchanged contact remains and only the deleted managed contact disappears from Outlook.
4. Delete the last contact in a selected book; the corresponding managed Outlook contact disappears.
5. Add a local Outlook contact without markers and keep contacts from a disabled book/another server. All remain after synchronization.
6. Simulate a failed or truncated DAV report; no contacts should be deleted. Retry successfully.
7. Repeat with default-folder and separate-folder destinations. Trigger manual sync while another run is active; the second run reports that synchronization is already running.

Live Outlook/Nextcloud acceptance requires the Windows installation and is separate from automated parser/build checks.
