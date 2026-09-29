# Privates Branding (IBP)

Das IBP-Logo wird **nicht** im Repository gespeichert. Der Ordner `branding/` im Projektstamm steht in
`.gitignore` und bleibt nur auf dem Build-Rechner.

| Datei | Verwendung | Format |
|---|---|---|
| `branding/ibp/source/*.png` | Originale (beliebige Größe) | PNG |
| `branding/ibp/header.png` | Banner in allen NC-Connector-Fenstern | PNG, 96 px hoch |
| `branding/ibp/app.png` | Fenster- und Ribbon-Symbol | PNG, 256 × 256 px |

Ablauf:

1. Originale nach `branding/ibp/source/` legen (Dateinamen wie in `assets/`, mindestens
   `header-solid-blue-164x48.png` und `app.png`).
2. `tools\branding\Prepare-Branding.ps1` ausführen: schneidet auf das Logo zu und skaliert.
3. `build.ps1` wie gewohnt. Sind `header.png` und `app.png` vorhanden, bettet der Build sie statt der
   öffentlichen Bilder ein; der Kopfbereich übernimmt die Hintergrundfarbe des Banners.

Ohne `branding/ibp/` baut das Projekt mit den öffentlichen Standardbildern. `assets/` (README-Bilder auf
GitHub) und `src/NcTalkOutlookAddIn/Resources/` bleiben unverändert.
