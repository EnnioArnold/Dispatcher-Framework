# Dispatcher Framework

Ein wiederverwendbares Dispatcher-Projekt für UiPath (Windows, VB.NET). Der Dispatcher liest
Quelldaten aus einer Excel-Datei und lädt jede Zeile als einzelnes Queue-Item in eine
Orchestrator-Queue. Die Verarbeitung der Queue-Items übernimmt ein separater Performer-Prozess,
der nicht Teil dieses Projekts ist.

Das Projekt ist als Framework gedacht: Für einen neuen Anwendungsfall muss in der Regel nur die
Konfigurationsdatei `Data\Config.xlsx` angepasst werden. Eingriffe in die Workflows selbst sind
nur nötig, wenn sich die Struktur der Quelldatei ändert (andere Spaltennamen, Daten beginnen
nicht in Zelle A1 usw.) oder andere Felder in die Queue geladen werden sollen. Alle Fälle sind
weiter unten Schritt für Schritt beschrieben (Fall 1 bis 6, inkl. Checkliste für eine komplett
neue Datei).

---

## Inhalt

1. [Funktionsweise](#funktionsweise)
2. [Projektstruktur](#projektstruktur)
3. [Voraussetzungen](#voraussetzungen)
4. [Schnellstart](#schnellstart)
5. [Konfiguration über Config.xlsx](#konfiguration-über-configxlsx)
6. [Anforderungen an die Quelldatei](#anforderungen-an-die-quelldatei)
7. [Anpassung an eine neue Quelldatei](#anpassung-an-eine-neue-quelldatei)
8. [Aufbau der Queue-Items](#aufbau-der-queue-items)
9. [Fehlerbehandlung](#fehlerbehandlung)
10. [Logging und Namenskonventionen](#logging-und-namenskonventionen)
11. [Workflows und Argumente im Detail](#workflows-und-argumente-im-detail)
12. [Troubleshooting](#troubleshooting)

---

## Funktionsweise

`Main.xaml` ist als State Machine mit vier Zuständen aufgebaut. Jeder Zustand ruft genau einen
Sub-Workflow auf und fängt dessen Fehler lokal ab. Eine aufgetretene Exception wird nicht sofort
geworfen, sondern in der Variable `systemException` zwischengespeichert und erst im Endzustand
erneut geworfen. So gibt es genau eine Stelle, an der der Job als Ganzes fehlschlägt.

```
                    ┌──────────────────┐
          ┌─────────│       Init       │  ruft InitAllSettings.xaml auf
          │         └────────┬─────────┘
          │ Fehler           │ Erfolg
          │         ┌────────▼─────────┐
          ├─────────│  Get Source Data │  ruft ReadSourceData.xaml auf
          │ Fehler  └────────┬─────────┘
          │ oder             │ Daten vorhanden
          │ keine Daten      │
          │         ┌────────▼─────────┐
          │         │  Populate Queue  │  ruft PopulateQueue.xaml auf
          │         └────────┬─────────┘
          │                  │ immer
          │         ┌────────▼─────────┐
          └────────▶│   End Process    │  loggt Ergebnis, wirft ggf. Exception
                    └──────────────────┘
```

Die Transition-Bedingungen im Einzelnen:

| Von | Nach | Bedingung |
|-----|------|-----------|
| Init | Get Source Data | `systemException Is Nothing` |
| Init | End Process | `systemException IsNot Nothing` |
| Get Source Data | Populate Queue | `systemException Is Nothing AndAlso Not keineDaten` |
| Get Source Data | End Process | `systemException IsNot Nothing OrElse keineDaten` |
| Populate Queue | End Process | keine Bedingung, immer |

Der Endzustand unterscheidet drei Ausgänge: Ist `systemException` gesetzt, wird der Fehler
geloggt und per `Throw` an den Orchestrator weitergereicht (Job = Failed). Ist die Quelldatei
leer (`keineDaten = True`), endet der Lauf geordnet und ohne Fehler. Andernfalls wird der
erfolgreiche Abschluss geloggt.

Variablen auf Ebene der State Machine:

| Name | Typ | Zweck |
|------|-----|-------|
| `konfiguration_dict` | `Dictionary(Of String, Object)` | wird von Init aus `Config.xlsx` befüllt |
| `quelldaten_dt` | `DataTable` | wird von Get Source Data befüllt |
| `keineDaten` | `Boolean` | `True`, wenn die Quelldatei keine Datenzeilen enthält |
| `systemException` | `System.Exception` | trägt eine aufgetretene Exception zum Endzustand |

## Projektstruktur

```
Dispatcher Framework\
├── Main.xaml                       Einstiegspunkt, State Machine (siehe oben)
├── InitAllSettings.xaml            liest Data\Config.xlsx in konfiguration_dict
├── ReadSourceData.xaml             liest die Quell-Excel-Datei, prüft auf leere Daten
├── PopulateQueue.xaml              ergänzt Zusatzspalten und lädt Zeile für Zeile in die Queue
├── Main.xaml.sequence-backup.txt   ältere lineare Version, nur als Referenz (kein Workflow)
├── project.json                    Abhängigkeiten und Projekteinstellungen
└── Data\
    ├── Config.xlsx                 zentrale Konfiguration (Blatt "Settings")
    └── README.md                   Kurzbeschreibung des Data-Ordners
```

## Voraussetzungen

- UiPath Studio (Windows-Projekt, VB.NET) mit den Paketen `UiPath.System.Activities` und
  `UiPath.Excel.Activities`. Beide sind in `project.json` eingetragen und werden beim Öffnen
  automatisch wiederhergestellt.
- Eine im Orchestrator angelegte Queue (Name frei wählbar, muss nur mit dem Wert `QueueName`
  in der Konfiguration übereinstimmen).
- Der Robot muss zur Laufzeit mit dem Orchestrator verbunden sein, sonst schlägt
  `Add Queue Item` fehl.

Hinweis zur Queue-Einstellung **Enforce Unique References**: Der Dispatcher setzt als Referenz
den Wert der Spalte `Dokumentlink` (Details unten). Ist die Queue mit eindeutigen Referenzen
angelegt und die Quelldatei enthält denselben Link zweimal, wirft der Orchestrator beim zweiten
Item einen Fehler und der Lauf schlägt fehl. Entweder die Quelldaten vorab bereinigen, die Queue
ohne Unique References anlegen oder die Referenz-Logik erweitern (siehe
[Referenz-Logik ändern](#referenz-logik-ändern)).

## Schnellstart

1. Projekt in UiPath Studio öffnen und Abhängigkeiten wiederherstellen lassen.
2. Im Orchestrator eine Queue anlegen, z. B. `Dispatcher_Main_Queue`.
3. `Data\Config.xlsx` öffnen (Blatt `Settings`) und die drei Werte eintragen:
   `SourceDataPath`, `SourceSheetName`, `QueueName`.
4. Sicherstellen, dass die Quelldatei eine Spaltenüberschrift `Dokumentlink` enthält
   (oder den Workflow anpassen, siehe unten).
5. `Main.xaml` per Run/Debug starten. Im Output-Panel erscheinen die Log-Meldungen
   `Start | Main | Dispatcher` bis `Abgeschlossen | Main | Dispatcher`.

Für den unbeaufsichtigten Betrieb wird das Projekt wie üblich publiziert und im Orchestrator
als Prozess eingerichtet.

## Konfiguration über Config.xlsx

`InitAllSettings.xaml` liest das Blatt `Settings` der Datei `Data\Config.xlsx`. Erwartet werden
zwei Spalten, `Name` und `Value`. Jede Zeile mit gefülltem `Name` landet als Schlüssel-Wert-Paar
im Dictionary `konfiguration_dict` und ist damit in allen Sub-Workflows verfügbar. Leere Zeilen
werden übersprungen.

Aktuell verwendet das Framework diese Schlüssel:

| Name | Beispiel-Value | Verwendet in | Bedeutung |
|------|----------------|--------------|-----------|
| `SourceDataPath` | `Data\Input\SourceData.xlsx` | ReadSourceData, PopulateQueue | Pfad zur Quell-Excel-Datei, relativ zum Projektordner oder absolut |
| `SourceSheetName` | `Sheet1` | ReadSourceData | Blattname innerhalb der Quelldatei |
| `QueueName` | `Dispatcher_Main_Queue` | PopulateQueue | Name der Ziel-Queue, zeichengenau wie im Orchestrator |

Zwei Punkte, die erfahrungsgemäß Fehler verursachen:

- `QueueName` muss exakt mit der Queue im Orchestrator übereinstimmen, sonst meldet
  `Add Queue Item`, dass die Queue nicht existiert.
- Relative Pfade in `SourceDataPath` beziehen sich auf das Arbeitsverzeichnis des Prozesses.
  In Studio ist das der Projektordner; im unbeaufsichtigten Betrieb ist ein absoluter Pfad
  (z. B. ein UNC-Pfad auf ein Netzlaufwerk) die robustere Wahl.

Zusätzliche eigene Schlüssel können jederzeit als weitere Zeilen ergänzt werden. Sie werden
automatisch mitgelesen und stehen in den Workflows über
`in_konfiguration_dict("MeinSchluessel").ToString` zur Verfügung, ohne dass am Einlese-Code
etwas geändert werden muss.

Pfad, Blattname und ein Asset-Name sind als Argumente von `InitAllSettings.xaml` mit
Standardwerten hinterlegt (`Data\Config.xlsx`, `Settings`, `Dispatcher_Config_JSON`) und können
beim Aufruf in `Main.xaml` überschrieben werden. Das Argument `in_assetName` sowie einige
Variablen in den Workflows (`assetJson`, `chunkGroesse`, `primaerschluesselSpalte` u. a.) sind
für spätere Erweiterungen reserviert (Orchestrator-Asset als Konfigurations-Override,
Bulk-Upload in Chunks) und derzeit ohne Funktion.

## Anforderungen an die Quelldatei

So erwartet der Dispatcher die Quelldatei im Auslieferungszustand:

```
     A                              B              C
 1   Dokumentlink                   Kunde          Betrag        ← Überschriften in Zeile 1
 2   \\server\scans\re_4711.pdf     Müller GmbH    1.250,00      ← erste Datenzeile
 3   \\server\scans\re_4712.pdf     Schulze AG       980,50
 ...
```

- Die Überschriften stehen in der ersten Zeile des benutzten Bereichs, die Daten direkt darunter.
  `ReadSourceData.xaml` liest den kompletten Bereich mit `Add Headers`; die Überschriften werden
  zu Spaltennamen der DataTable.
- Eine Spalte muss exakt `Dokumentlink` heißen. Ihr Wert wird (getrimmt) als Queue-Referenz
  verwendet und zusätzlich als Feld `dokumentLink` ins Queue-Item geschrieben. Fehlt die Spalte,
  bricht `PopulateQueue.xaml` mit einer Exception ab.
- Alle weiteren Spalten dürfen beliebig heißen. Sie werden derzeit nicht ins Queue-Item
  übernommen; wie das geht, steht unter
  [Weitere Spalten ins Queue-Item übernehmen](#weitere-spalten-ins-queue-item-übernehmen).
- Die Spaltennamen `Reference`, `createTS`, `Filename` und `ExcelReihe` sind reserviert, weil
  `PopulateQueue.xaml` sie selbst anlegt. Kommen sie bereits in der Quelldatei vor, schlägt das
  Anlegen der Spalte fehl.

Ist die Datei leer (keine Datenzeilen unterhalb der Überschriften), endet der Lauf geordnet mit
dem Log `Keine Quelldaten | Main | Dispatcher wird ohne Fehler beendet`.

## Anpassung an eine neue Quelldatei

Hier die typischen Fälle, sortiert von "nur Konfiguration" bis "Workflow anpassen".

### Fall 1: Andere Datei oder anderes Blatt, gleiche Struktur

Es reicht, in `Data\Config.xlsx` die Werte `SourceDataPath` und `SourceSheetName` zu ändern.
Kein Eingriff in die Workflows nötig.

### Fall 2: Daten beginnen nicht in Zelle A1

Beispiel: Die Überschriften stehen erst ab Spalte E (also der fünften Spalte), oder oberhalb der
Tabelle liegen Titel- und Leerzeilen:

```
     A       B       C       D       E              F           G
 1                                   (leer)
 2                   Quartalsreport 2026
 3
 4                                   Dokumentlink   Kunde       Betrag     ← Überschriften
 5                                   \\srv\a.pdf    Müller      1.250,00   ← erste Datenzeile
```

`ReadSourceData.xaml` liest den Bereich aktuell mit einem leeren Range (`""`), also den gesamten
benutzten Bereich des Blatts. Bei vorgelagerten Titelzeilen oder Leerspalten würden dadurch
falsche Überschriften eingelesen. Die Lösung ist, den Startpunkt explizit anzugeben:

1. `ReadSourceData.xaml` in Studio öffnen.
2. Die Aktivität **"Read Source Excel"** (Workbook Read Range) auswählen.
3. Im Eigenschaften-Panel die Eigenschaft **Range** von `""` auf die Zelle setzen, in der die
   **erste Überschrift** steht — im Beispiel oben `"E4"`.

Wird nur eine Startzelle angegeben, liest UiPath von dort bis zum Ende der zusammenhängenden
Daten. Alternativ kann ein fester Bereich wie `"E4:G500"` angegeben werden, wenn rechts oder
unterhalb der Tabelle weitere Inhalte liegen, die nicht mitgelesen werden sollen.

Wer den Startpunkt lieber konfigurierbar hält, statt ihn im Workflow zu verdrahten: in
`Config.xlsx` eine Zeile `SourceRange` / `E4` ergänzen und im Range-Feld der Aktivität
`in_konfiguration_dict("SourceRange").ToString` eintragen. Dann lässt sich der Bereich künftig
ohne Studio ändern.

Wichtig für das Feld `excelRow`: Der Zeilenzähler in `PopulateQueue.xaml` startet bei 2, geht
also davon aus, dass die Überschriften in Zeile 1 stehen. Beginnen die Daten wie im Beispiel
erst in Zeile 5, muss der Default der Variable `zeilenZaehler` in `PopulateQueue.xaml` von `2`
auf `5` gesetzt werden, damit die Zeilennummern im Queue-Item wieder zur echten Excel-Datei
passen.

### Fall 3: Die Schlüsselspalte heißt nicht "Dokumentlink"

Beispiel: Die neue Datei hat statt `Dokumentlink` eine Spalte `Rechnungsnummer`, die als
Referenz dienen soll. Der Spaltenname ist an drei Stellen in `PopulateQueue.xaml` als
String-Literal hinterlegt:

| Stelle | Aktivität | Ausdruck ändern |
|--------|-----------|-----------------|
| Referenz setzen | Schleife "Zusatzfelder je Zeile setzen", erstes Assign (`Reference`) | `zeile("Dokumentlink").ToString.Trim` → `zeile("Rechnungsnummer").ToString.Trim` |
| Queue-Item-Feld | Aktivität "Add Queue Item", ItemInformation-Eintrag `dokumentLink` | `zeile("Dokumentlink").ToString` → `zeile("Rechnungsnummer").ToString` (bei Bedarf auch den Feldnamen `dokumentLink` umbenennen) |
| Erfolgs-Log | "Log Message" nach dem Upload | `zeile("Dokumentlink").ToString` → `zeile("Rechnungsnummer").ToString` |

Nach der Änderung einmal mit einer kleinen Testdatei laufen lassen und im Orchestrator prüfen,
ob Referenz und Felder wie erwartet gefüllt sind.

### Fall 4: Weitere Spalten ins Queue-Item übernehmen

Standardmäßig landen nur die vier unten beschriebenen Felder im Queue-Item. Sollen zusätzliche
Spalten der Quelldatei mitgegeben werden (z. B. `Kunde` und `Betrag`):

1. `PopulateQueue.xaml` öffnen und die Aktivität **"Add Queue Item"** in der Schleife
   "For Each: Einträge einzeln hochladen" auswählen.
2. Im Eigenschaften-Panel unter **ItemInformation** neue Einträge ergänzen, z. B.
   Feldname `kunde` mit Wert `zeile("Kunde").ToString`.

Die Spalten müssen dafür nicht vorher angelegt werden; sie kommen ja bereits aus der Quelldatei.
Ein Eintrag in "Spalten anlegen" ist nur für berechnete Felder nötig, die es in der Quelle nicht
gibt (nach dem Muster der vorhandenen Spalten `createTS`, `Filename`, `ExcelReihe`).

### Fall 5: Referenz-Logik ändern

Aktuell ist die Referenz der reine `Dokumentlink`-Wert. Soll derselbe Datensatz pro Tag nur
einmal in die Queue gelangen (klassisches Dispatcher-Muster mit Enforce Unique References),
hängt man in der Schleife "Zusatzfelder je Zeile setzen" im Assign für `Reference` einen
Datumsstempel an:

```vb
zeile("Dokumentlink").ToString.Trim & "_" & DateTime.Now.ToString("yyyyMMdd")
```

Für mehrere Läufe pro Tag entsprechend `yyyyMMddHHmmss`. Zu beachten: Die Referenz ist im
Orchestrator auf 128 Zeichen begrenzt. Bei langen UNC-Pfaden oder URLs als Basis empfiehlt sich
stattdessen eine kürzere Schlüsselspalte oder ein Hash.

### Fall 6: Komplett neue Datei mit anderen Feldern — Checkliste

Der häufigste Praxisfall: Eine neue Quelldatei mit ganz anderen Spalten soll in eine (neue)
Queue geladen werden, und statt der bisherigen Felder sollen andere Werte ins Queue-Item.
Beispiel: Die neue Datei hat die Spalten `Auftragsnummer`, `Kunde`, `Lieferdatum` — kein
`Dokumentlink` mehr.

Dann sind der Reihe nach diese Stellen zu ändern (Fall 1 bis 4 kombiniert):

| Nr. | Wo | Was |
|-----|----|----|
| 1 | `Data\Config.xlsx` | `SourceDataPath`, `SourceSheetName` und ggf. `QueueName` auf die neue Datei/Queue setzen |
| 2 | `ReadSourceData.xaml`, Aktivität "Read Source Excel" | Nur falls die Daten nicht in A1 beginnen: Range auf die erste Überschriften-Zelle setzen (siehe Fall 2) |
| 3 | `PopulateQueue.xaml`, Schleife "Zusatzfelder je Zeile setzen", Assign `Reference` | Referenz auf die neue Schlüsselspalte umstellen, z. B. `zeile("Auftragsnummer").ToString.Trim` |
| 4 | `PopulateQueue.xaml`, Aktivität "Add Queue Item", **ItemInformation** | Alte Feld-Einträge, die es in der neuen Datei nicht gibt, **entfernen oder umstellen** (v. a. `dokumentLink` → z. B. `auftragsnummer`) und für jede gewünschte neue Spalte einen Eintrag ergänzen, z. B. `kunde` = `zeile("Kunde").ToString`, `lieferdatum` = `zeile("Lieferdatum").ToString` |
| 5 | `PopulateQueue.xaml`, "Log Message" nach dem Upload | Den Ausdruck `zeile("Dokumentlink").ToString` durch die neue Schlüsselspalte ersetzen |
| 6 | `PopulateQueue.xaml`, Variable `zeilenZaehler` | Nur falls die Überschriften nicht in Zeile 1 stehen: Default anpassen (siehe Fall 2) |

Wichtig bei Schritt 4: Jeder ItemInformation-Eintrag, der auf eine Spalte verweist, die es in
der neuen Datei nicht gibt, führt zur Laufzeit zu einem Abbruch mit
"Column '...' does not belong to table". Deshalb immer **alle** vorhandenen Einträge
durchgehen, nicht nur neue ergänzen. Die vom Framework selbst erzeugten Felder `createTS`,
`filename` und `excelRow` funktionieren dagegen mit jeder Datei unverändert und können bleiben.

Nicht ändern muss man: `Main.xaml`, `InitAllSettings.xaml` und die Spalten-Anlage am Anfang von
`PopulateQueue.xaml` ("Spalten anlegen") — letztere nur dann, wenn zusätzliche **berechnete**
Felder gebraucht werden, die es in der Quelldatei nicht gibt.

Danach mit einer kleinen Testdatei (2–3 Zeilen) laufen lassen und im Orchestrator ein
Queue-Item öffnen: Unter "Specific Data" müssen alle neuen Felder mit den erwarteten Werten
stehen, die Referenz muss dem neuen Schlüssel entsprechen.

## Aufbau der Queue-Items

Pro Datenzeile entsteht ein Queue-Item mit Priorität `Normal`. `PopulateQueue.xaml` legt dafür
zu Beginn vier Hilfsspalten in der DataTable an und befüllt sie je Zeile:

| Queue-Feld | Quelle | Beispiel |
|------------|--------|----------|
| Reference (Item-Referenz) | Spalte `Dokumentlink`, getrimmt | `\\server\scans\re_4711.pdf` |
| `createTS` | Zeitstempel beim Befüllen (`yyyy-MM-dd HH:mm:ss`), pro Zeile gesetzt | `2026-07-14 08:15:02` |
| `filename` | Dateiname der Quell-Excel-Datei, per `Path.GetFileName` aus `SourceDataPath` | `SourceData.xlsx` |
| `dokumentLink` | Spalte `Dokumentlink`, unverändert | `\\server\scans\re_4711.pdf` |
| `excelRow` | Laufende Zeilennummer in der Quelldatei (Header = Zeile 1, erste Datenzeile = 2) | `2` |

Das Feld `excelRow` ist für den Performer gedacht: Damit lässt sich ein Ergebnis oder ein
Fehlerstatus später exakt in die richtige Zeile der Quelldatei zurückschreiben.

## Fehlerbehandlung

- Jeder der drei Sub-Workflows wird in `Main.xaml` in einem eigenen Try/Catch aufgerufen. Im
  Catch wird der Fehler mit dem Namen des betroffenen Workflows geloggt
  (`Fehlgeschlagen | <Workflowname> | <Meldung + StackTrace>`) und in `systemException` abgelegt.
- Der Endzustand wirft `systemException` erneut. Der Job erscheint im Orchestrator als Failed,
  mit der ursprünglichen Fehlermeldung.
- Eine leere Quelldatei ist kein Fehler: `ReadSourceData.xaml` setzt `out_keineDaten = True`,
  die State Machine geht direkt zu End Process und der Job endet Successful.
- Der Upload läuft Zeile für Zeile ohne eigenes Try/Catch pro Item. Schlägt ein einzelnes
  `Add Queue Item` fehl (z. B. doppelte Referenz bei Enforce Unique References, Queue nicht
  gefunden, Verbindung weg), bricht der gesamte Lauf ab — bereits hochgeladene Items bleiben
  dabei in der Queue. Wer stattdessen "loggen und weitermachen" möchte, legt um das
  `Add Queue Item` in `PopulateQueue.xaml` ein Try/Catch und protokolliert den Fehler als
  Warnung.

## Logging und Namenskonventionen

Das Projekt folgt den DSD-Conventions für RPA.

Log-Meldungen sind einheitlich mit Pipe-Trennern aufgebaut und verwenden nur die Level Info,
Warn und Error:

```
Start         | <Workflowname> | <Zusatzinfo>
Abgeschlossen | <Workflowname> | <Kennzahl>=<Wert>
Fehlgeschlagen| <Workflowname> | <Fehlertext>
```

Beispiele aus einem realen Lauf:

```
Start | Main | Dispatcher
Start | InitAllSettings
Abgeschlossen | InitAllSettings | AnzahlSchluessel=3
Start | ReadSourceData
Abgeschlossen | ReadSourceData | Zeilen=125
Start | PopulateQueue | Queue=Dispatcher_Main_Queue
Abgeschlossen | UploadQueue | Link: \\server\scans\re_4711.pdf | Zeile: 2
...
Abgeschlossen | Main | Dispatcher
```

Variablen und Argumente sind auf Deutsch benannt (technische Begriffe wie Queue oder Asset
bleiben englisch), in lower camelCase. Argumente tragen ein Richtungs-Präfix (`in_`, `out_`,
`io_`), Sammeltypen ein Typ-Suffix: `_dt` für DataTable, `_dict` für Dictionary, `_liste` für
List. Beispiele: `in_quelldaten_dt`, `io_konfiguration_dict`, `quellPfad`.

Hinweis zum Workflow Analyzer: Die UiPath-Standardregeln erwarten ein `dt_`-Präfix statt der
DSD-Suffixe und melden deshalb Namens-Warnungen. Das ist bekannt und bewusst so belassen, bis
ein DSD-eigenes Analyzer-Ruleset in der Studio-Governance hinterlegt ist.

## Workflows und Argumente im Detail

### InitAllSettings.xaml

Liest `Config.xlsx` und befüllt das Konfigurations-Dictionary. Ist das übergebene Dictionary
`Nothing`, wird ein neues angelegt; vorhandene Einträge bleiben ansonsten erhalten und werden
bei gleichem Schlüssel überschrieben.

| Argument | Richtung | Typ | Default |
|----------|----------|-----|---------|
| `io_konfiguration_dict` | In/Out | `Dictionary(Of String, Object)` | – |
| `in_konfigurationsPfad` | In | `String` | `Data\Config.xlsx` |
| `in_konfigurationsBlatt` | In | `String` | `Settings` |
| `in_assetName` | In | `String` | `Dispatcher_Config_JSON` (reserviert, derzeit ungenutzt) |

### ReadSourceData.xaml

Liest die Quell-Excel-Datei (Pfad und Blatt aus der Konfiguration) per Workbook Read Range mit
`Add Headers` und prüft anschließend, ob Datenzeilen vorhanden sind.

| Argument | Richtung | Typ |
|----------|----------|-----|
| `in_konfiguration_dict` | In | `Dictionary(Of String, Object)` |
| `out_quelldaten_dt` | Out | `DataTable` |
| `out_keineDaten` | Out | `Boolean` |

### PopulateQueue.xaml

Legt die Zusatzspalten `Reference`, `createTS`, `Filename`, `ExcelReihe` an, befüllt sie je
Zeile und lädt anschließend jede Zeile per `Add Queue Item` in die Queue.

| Argument | Richtung | Typ |
|----------|----------|-----|
| `in_quelldaten_dt` | In | `DataTable` |
| `in_konfiguration_dict` | In | `Dictionary(Of String, Object)` |

## Troubleshooting

| Symptom | Wahrscheinliche Ursache | Lösung |
|---------|------------------------|--------|
| `Fehlgeschlagen \| InitAllSettings \| ...` gleich zu Beginn | `Data\Config.xlsx` fehlt, ist geöffnet/gesperrt oder das Blatt heißt nicht `Settings` | Datei und Blattname prüfen, Excel schließen |
| `Fehlgeschlagen \| ReadSourceData \| ...` | `SourceDataPath` falsch, Datei gesperrt oder `SourceSheetName` existiert nicht | Pfad absolut angeben, Blattnamen zeichengenau prüfen |
| Fehler "Column 'Dokumentlink' does not belong to table" | Quelldatei hat keine Spalte `Dokumentlink` bzw. Überschrift weicht ab (Leerzeichen, Tippfehler) | Überschrift korrigieren oder Workflow anpassen (Fall 3) |
| Fehler "Queue ... does not exist" beim Upload | `QueueName` stimmt nicht mit dem Orchestrator überein oder der Robot läuft im falschen Ordner | Queue-Name und Orchestrator-Ordner prüfen |
| Fehler wegen doppelter Referenz beim Upload | Queue hat Enforce Unique References und `Dokumentlink` kommt mehrfach vor (oder zweiter Lauf am selben Tag mit Datums-Referenz) | Quelldaten bereinigen oder Referenz-Logik erweitern (Fall 5) |
| Spaltennamen in der DataTable sind `Column1`, `Column2`, ... | Über den Überschriften liegen Titel-/Leerzeilen, der Read Range beginnt zu früh | Range in `ReadSourceData.xaml` auf die erste Überschriften-Zelle setzen (Fall 2) |
| `excelRow` passt nicht zur echten Zeilennummer | Überschriften stehen nicht in Zeile 1 | Default von `zeilenZaehler` in `PopulateQueue.xaml` anpassen (Fall 2) |
| Job endet Successful, aber Queue bleibt leer | Quelldatei war leer (`Keine Quelldaten` im Log) | Quelldatei bzw. `SourceDataPath`/Range prüfen |

## Hinweis zur lokalen Kommandozeile

Das Projekt wurde mit der `uip`-CLI erstellt und geprüft. Für `uip rpa validate` bzw.
`uip rpa build` auf der Kommandozeile wird das .NET 8 SDK auf dem PATH benötigt; es wurde
bereits benutzerlokal unter `%LOCALAPPDATA%\Microsoft\dotnet` installiert. Für die Arbeit in
UiPath Studio selbst ist davon nichts nötig.
