# My Nextcloud in Outlook

The Outlook flow should match the Thunderbird sharing window wherever native Outlook controls allow it. This work covers local files and the user's own Nextcloud. Other VFS providers are not part of it.

- [x] Start the queue step with the same “Build share” heading and a clearly labelled destination folder. Show the complete destination path below it, not only the configured base folder.
- [x] Use the same visible wording as Thunderbird for the source buttons, queue summary, source headings, file status and footer buttons in every language. Outlook only leaves out “Other source”.
- [x] Replace the current flat file list with a queue split into “Local” and “My Nextcloud”. Keep the `+ Local` and `+ My Nextcloud` source buttons above it, each with file and folder choices.
- [x] Show folders as an expandable tree. Files need matching icons, their size and status, and a remove button.
- [x] Match Thunderbird's summary row: entries, sources and total size on the left; storage limit and occupied Nextcloud storage on the right.
- [x] Match the queue states shown in Thunderbird: selected folders start expanded, files show “Waiting” before upload, and Back, Upload, Next and Cancel follow the same enabled states.
- [x] Build a Nextcloud file and folder picker. Match Thunderbird's search, name, size, modified date, original Nextcloud and file icons, image preview and storage display. Use one compact Explorer-style bar with back, forward, up and refresh controls followed by the Nextcloud root symbol and clickable path segments; do not repeat the account name or add a second root symbol. Report the item count for multiple selections instead of “No file selected”.
- [x] Copy selected Nextcloud files and folders into the share on the server. Leave the originals untouched and do not download them first.
- [x] Handle local and Nextcloud items in the same share. Duplicate names, empty folders, progress, cancellation and cleanup must keep working.
- [x] Keep attachment automation on the same queue.
- [x] Add every new text to all existing translations.
- [x] Check keyboard use and both light and dark Windows themes.
- [x] Update the user and development documentation.
- [x] Run the tests and build, then verify that the work folder and feature branch match.
