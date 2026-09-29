# Outlook CalDav Synchronizer (mitgeliefert)

| | |
|---|---|
| Version | 4.7.1 (2026-08-17) |
| Datei | `CalDavSynchronizer.Setup.msi` aus `OutlookCalDavSynchronizer-4.7.1.zip`, unverändert |
| SHA-256 (MSI) | `9624D688C832286683D40BF7B946C9C18228BE2AC0D89C8772A20E91520F0FDA` |
| SHA-256 (ZIP) | `46AE790A0518936AD1E58D3C6723F53497136D35E88B1AF56C1DFE299686846C` |
| Download | https://github.com/aluxnimm/outlookcaldavsynchronizer/releases/tag/v4.7.1 |
| Quellcode | https://github.com/aluxnimm/outlookcaldavsynchronizer |
| Lizenz | GNU Affero General Public License v3.0, Text in `LICENSE.txt` (aus dem Upstream-Tag v4.7.1) |
| Voraussetzungen | .NET Framework 4.8, Visual Studio Tools for Office Runtime (in Office 2016 / Microsoft 365 enthalten) |

Das Paket wird nicht verändert. NC Connector richtet beim ersten Outlook-Start das Kalender-Profil
"Nextcloud Kalender" für den persönlichen Nextcloud-Kalender ein (Outlook-Standardkalender, beidseitig).
Vorhandene Profile der Benutzer werden nicht verändert. Das Profil wird nach einem Outlook-Neustart aktiv;
NC Connector weist die Benutzer darauf hin.

## Verteilung per Gruppenrichtlinie

Der Ordner `GPO-<Version>` aus `build.ps1` enthält:

- `CalDavSynchronizer.Setup.msi` und `CalDavSynchronizer-AllUsers.mst`
- `IBP-NC-ConnectorForOutlook<Version>.msi`

Beide MSI-Dateien auf eine Freigabe legen, die Benutzer- bzw. Computerkonten lesen dürfen. Die Reihenfolge
der Pakete ist egal: NC Connector richtet den Kalender ein, sobald CalDav Synchronizer in Outlook vorhanden
ist.

### Variante A: Benutzerkonfiguration (wie bisher für NC Connector)

CalDav Synchronizer wird pro Benutzer installiert, wie bei einer manuellen Installation "nur für mich".

1. Benutzerkonfiguration → Richtlinien → Softwareeinstellungen → Softwareinstallation → Neu → Paket →
   `CalDavSynchronizer.Setup.msi` über den Freigabepfad (UNC) auswählen → **Zugewiesen**.
   **Ohne** `CalDavSynchronizer-AllUsers.mst` — die Transformation ist nur für Variante B.
2. Paket-Eigenschaften → Registerkarte *Bereitstellung* → **"Diese Anwendung bei der Anmeldung installieren"**
   aktivieren. Ohne diese Option wird das Add-in nur angekündigt und nie installiert, weil es keine
   Verknüpfung gibt, die die Installation auslöst.
3. `IBP-NC-ConnectorForOutlook<Version>.msi` wie bisher zuweisen.

Hinweise:

- Mehrere Benutzer auf einem Rechner erhalten jeweils eine eigene Installation.
- Benutzer, die CalDav Synchronizer schon selbst installiert haben, zuerst mit einem Testkonto prüfen:
  Windows Installer übernimmt die vorhandene Installation in der Regel in die Verwaltung durch die GPO.
- Für eine neue CalDav-Synchronizer-Version ein neues Paket anlegen und auf der Registerkarte
  *Aktualisierungen* das bisherige Paket als zu aktualisierendes Paket eintragen.

### Variante B: Computerkonfiguration

CalDav Synchronizer wird einmal pro Rechner für alle Benutzer installiert.

1. Computerkonfiguration → Richtlinien → Softwareeinstellungen → Softwareinstallation → Neues Paket →
   `CalDavSynchronizer.Setup.msi` → **Erweitert**
   → Registerkarte *Änderungen* → `CalDavSynchronizer-AllUsers.mst` hinzufügen. Die Transformation setzt
   `ALLUSERS=1`, damit das Add-in für alle Benutzer des Rechners registriert wird. Sie kann nur beim Anlegen
   des Pakets hinzugefügt werden.
2. `IBP-NC-ConnectorForOutlook<Version>.msi` zuweisen.

Nicht beide Varianten für dieselben Benutzer mischen, sonst ist CalDav Synchronizer doppelt registriert.

## Manuell / Testgruppe

`IBP-NC-Connector-Setup-<Version>.exe` installiert beide Pakete in einem Schritt (still: `/quiet`).
