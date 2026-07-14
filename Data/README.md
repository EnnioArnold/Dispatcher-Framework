# Data — Konfiguration & Quelldaten

Dieses Verzeichnis enthält die Laufzeit-Artefakte des Dispatcher-Frameworks. Die Dateien werden
**nur zur Laufzeit** benötigt (nicht für `validate`/`build`).

## `Config.xlsx`

Excel-Arbeitsmappe mit dem Blatt **`Settings`**, zwei Spalten: **`Name`** und **`Value`**.
Wird von `InitAllSettings.xaml` gelesen. Erwartete Schlüssel:

| Name | Beispiel-Value | Beschreibung |
|------|----------------|--------------|
| `SourceDataPath` | `Data\Input\SourceData.xlsx` | Pfad zur Quell-Excel-Datei |
| `SourceSheetName` | `Sheet1` | Blattname der Quelldaten |
| `QueueName` | `Dispatcher_Main_Queue` | Ziel-Queue im Orchestrator |
| `PrimaryKeyColumn` | `EditierteKundenNr` | Spalte, die den eindeutigen Primärschlüssel liefert |
| `ChunkSize` | `10000` | Max. Zeilen pro Bulk-Upload (≤ 10.000, optional; Default 10000) |

> `InitAllSettings.xaml` besitzt Standardwerte für Pfad (`Data\Config.xlsx`), Asset-Name
> (`Dispatcher_Config_JSON`) und Blattname (`Settings`). Diese können beim Aufruf überschrieben werden.

## Orchestrator-Text-Asset (moderner Override)

Zusätzlich liest `InitAllSettings.xaml` ein **Text-Asset** (Default-Name `Dispatcher_Config_JSON`)
und deserialisiert dessen JSON direkt in das `Config`-Dictionary. **JSON-Werte überschreiben die
Excel-Werte** — so lässt sich die Konfiguration zentral im Orchestrator pflegen, ohne die lokale
`Config.xlsx` anzufassen.

Beispiel-JSON-Inhalt des Assets:

```json
{
  "QueueName": "Dispatcher_Main_Queue",
  "PrimaryKeyColumn": "EditierteKundenNr",
  "ChunkSize": 5000
}
```

## Queue-Voraussetzung im Orchestrator

Die Ziel-Queue muss mit **„Enforce Unique References = Yes"** angelegt sein. `PopulateQueue.xaml`
erzeugt vor dem Upload eine `Reference`-Spalte (`<Primärschlüssel>_<yyyyMMdd>`); bereits am selben
Tag dispatchte Schlüssel werden vom Orchestrator als Duplikat (Fehler **1016**) abgewiesen und
vom Framework übersprungen, ohne den Lauf abzubrechen.

## Zusätzliche Spalten, die PopulateQueue.xaml selbst erzeugt

Neben `Reference` fügt `PopulateQueue.xaml` vor dem Upload drei weitere Spalten in die Quelldaten
ein, die dann als eigene Felder im Queue-Item landen:

| Spalte | Typ | Inhalt |
|--------|-----|--------|
| `createTS` | String | Zeitpunkt dieses Dispatcher-Laufs (`yyyy-MM-dd HH:mm:ss`), für alle Zeilen desselben Laufs identisch |
| `Filename` | String | Dateiname, extrahiert aus der Spalte `Dokumentlink` (funktioniert für lokale Pfade, UNC-Pfade und HTTP(S)-Links) |
| `ExcelReihe` | Int32 | Tatsächliche Zeilennummer in der Quell-Excel-Datei (Zeile 1 = Header, erste Datenzeile = Zeile 2) |

Diese Spaltennamen sind im Workflow als String-Literale hinterlegt (je ein `If`/`Assign`-Paar am
Anfang von `PopulateQueue.xaml` sowie vier `Assign`-Aktivitäten in der Schleife "Zusatzfelder je
Zeile setzen"). Möchtest du andere Feldnamen, änderst du dort die Strings.
