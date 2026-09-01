# New-UiPathDispatcher.ps1

Ein Assistent, der aus einem bestehenden UiPath-Dispatcher-Template ein neues,
fertig konfiguriertes Dispatcher-Projekt macht. Statt in Studio 40 Assign-Zeilen
von Hand umzubauen, wählt man das Quell-Excel aus, hakt die gewünschten Spalten
an, baut die Queue-Referenz zusammen und drückt auf Erstellen.

Das Skript läuft lokal, braucht kein installiertes Excel, keine PowerShell-Module
und keine Orchestrator-Verbindung. Es liest und schreibt `.xlsx` direkt über
OpenXML (ZIP + XML) und bearbeitet die `.xaml`-Dateien über das DOM, nicht über
Textersetzung.

---

## Inhalt

1. [Wozu das Ganze](#wozu-das-ganze)
2. [Voraussetzungen](#voraussetzungen)
3. [Schnellstart](#schnellstart)
4. [Das Template, das erwartet wird](#das-template-das-erwartet-wird)
5. [Was der Generator konkret ändert](#was-der-generator-konkret-ändert)
6. [Die fünf Masken](#die-fünf-masken)
7. [**Queue-Item-Felder im Detail**](#queue-item-felder-im-detail) ← der wichtigste Abschnitt
8. [Die Queue-Referenz](#die-queue-referenz)
9. [Konfiguration DEV / PRD](#konfiguration-dev--prd)
10. [Nach dem Generieren: Checkliste in Studio](#nach-dem-generieren-checkliste-in-studio)
11. [Grenzen und bekannte Eigenheiten](#grenzen-und-bekannte-eigenheiten)
12. [Fehlersuche](#fehlersuche)
13. [Wie es intern funktioniert](#wie-es-intern-funktioniert)

---

## Wozu das Ganze

Ein Dispatcher ist immer dasselbe: Excel einlesen, pro Zeile ein Queue-Item
bauen, hochladen. Was sich von Prozess zu Prozess unterscheidet, sind genau drei
Dinge — welche Spalten mitgehen, unter welchen Schlüsseln sie im Queue-Item
landen, und wie die Reference aussieht.

Genau diese drei Dinge sind im Template aber über mehrere Stellen verteilt:

```
Utility_Files/PopulateQueue.xaml
   ├── Sequence.Variables ............ eine String-Variable pro Spalte
   ├── MultipleAssign ................ CurrentRow("Spalte") -> Variable
   ├── Assign "QItemInfo" ............ Dictionary-Literal mit allen Schlüsseln
   └── AddQueueItem .................. Reference, QueueType, FolderPath
```

Vier Stellen, die zueinander passen müssen. Wer eine Spalte vergisst, merkt es
erst zur Laufzeit. Der Generator schreibt alle vier aus einer einzigen Liste,
damit sie gar nicht erst auseinanderlaufen können.

---

## Voraussetzungen

| Was | Warum |
|---|---|
| Windows mit PowerShell 5.1 oder neuer | WPF-Oberfläche |
| Ein Dispatcher-Template als Ordner | wird kopiert, nie verändert |
| Ein Quell-Excel im OpenXML-Format (`.xlsx` oder `.xlsm`) | daraus kommen Sheets und Spaltennamen |

Nicht nötig: Excel, UiPath Studio, Orchestrator-Zugriff, Admin-Rechte,
Internetverbindung, irgendein PowerShell-Modul.

Das Skript braucht einen STA-Thread. Falls es in einem MTA-Host gestartet wird
(passiert z. B. in manchen IDE-Terminals), startet es sich selbst mit `-STA` neu.
Man muss sich darum nicht kümmern.

Wenn die Hausschrift *Futura Now Text* nicht installiert ist, greift die
Oberfläche automatisch auf Futura → Century Gothic → Segoe UI zurück. Sieht dann
etwas anders aus, funktioniert aber identisch.

---

## Schnellstart

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\New-UiPathDispatcher.ps1
```

Wer die Pfade schon kennt, spart sich zwei Klicks:

```powershell
.\New-UiPathDispatcher.ps1 `
    -TemplatePath "C:\RPA\Templates\Dispatcher_Vorlage" `
    -OutputPath   "C:\RPA\Projekte\Zahlungen_Dispatcher"
```

Beide Parameter sind optional. `-OutputPath` füllt nur das Feld der ersten
Maske vor. `-TemplatePath` liest zusätzlich `Data\Config_DEV.xlsx` und
`Config_PRD.xlsx` des Templates ein und befüllt damit die Konfigurations-
tabelle auf Maske 4 — genau wie das Wählen des Ordners über *Durchsuchen*.

Danach in dieser Reihenfolge:

```
   1              2              3              4              5
┌────────┐   ┌────────┐   ┌────────┐   ┌────────┐   ┌────────┐
│Vorlage │──▶│ Felder │──▶│Referenz│──▶│ Config │──▶│Bestäti-│──▶ Erstellen
│& Quelle│   │        │   │        │   │DEV/PRD │   │ gung   │
└────────┘   └────────┘   └────────┘   └────────┘   └────────┘
 Template     Spalten      Reference    SourceData   Prüfen und
 Ziel         an/aus       zusammen-    Path, Queue  bestätigen
 Excel        Schlüssel    stecken      Name, ...
 Sheet        Reihenfolge
```

Zwischen den Masken kann man beliebig hin- und herspringen. Beim Klick auf
*Weiter* wird die aktuelle Maske geprüft; über die Reiter oben kommt man auch
ohne Prüfung weiter. Umgehen lässt sich die Prüfung damit aber nicht: *Dispatcher
erstellen* prüft die Masken 1–4 noch einmal alle nach und springt zur ersten
unvollständigen zurück.

---

## Das Template, das erwartet wird

```
Dispatcher_Vorlage\
├── Main.xaml                        StateMachine, Argument In_Umgebung
├── project.json
├── Data\
│   ├── Config_DEV.xlsx              Sheet "Settings", Spalten Name | Value
│   └── Config_PRD.xlsx
└── Utility_Files\                   Ordnername egal, wird erkannt
    ├── InitAllSettings.xaml         Config -> konfiguration_dict
    ├── ReadSourceData.xaml          Excel -> quelldaten_dt
    └── PopulateQueue.xaml           DataRow -> Queue Item
```

Gefunden wird das Template über `PopulateQueue.xaml`. Heisst der Utility-Ordner
`Utility Files`, `Utility_Files` oder `Framework` — egal, der Generator sucht die
Datei rekursiv und passt die Invoke-Pfade in `Main.xaml` an den tatsächlichen
Ordnernamen an.

Der Ablauf im generierten Projekt:

```
Main.xaml  (StateMachine "Dispatcher Framework")

  ┌──────┐  In_Umgebung        ┌────────────────┐        ┌───────────────┐
  │ Init │  SZT1/SZT2 -> DEV   │ Get Source Data│        │ Populate Queue│
  │      │─▶SZP       -> PRD  ─▶│                │───────▶│               │
  └──────┘                     └────────────────┘  Daten └───────────────┘
      │                                │            da?          │
      │ InitAllSettings                │ ReadSourceData          │ PopulateQueue
      ▼                                ▼                         ▼
  Data\Config_{DEV|PRD}.xlsx    SourceDataPath\*.xlsx     pro DataRow ein
  Sheet "Settings"              Sheet = SourceSheetName    AddQueueItem
  Name/Value -> konfiguration_dict   -> quelldaten_dt
                                     -> dokumentname
                                            │
                                     keine Daten? ──▶ End Process (ohne Fehler)
```

---

## Was der Generator konkret ändert

Kopiert wird alles ausser `.local`, `.settings`, `.objects`, `.entities`,
`.tmh`, `.git`, `.vs`, `.screenshots`, `.DS_Store` — und ausser dem Generator
selbst, falls er im Template-Ordner liegt.

### project.json

```diff
- "name": "Versandregeln Dispatcher",
+ "name": "Zahlungen Dispatcher",
- "projectId": "d9580892-2612-44ac-b4fc-adfb0d8a9109",
+ "projectId": "<neue GUID>",
- "description": "Leerer Prozess",
+ "description": "<Text aus Maske 1>",
```

Auch die `uniqueId` der Entry Points wird neu vergeben. Ohne das hätte man zwei
Projekte mit identischer Projekt-GUID im Orchestrator — das gibt später
Verwechslungen beim Publizieren.

### Main.xaml

```diff
- this:Main.In_Umgebung="SZT2"
+ this:Main.In_Umgebung="SZT1"

- WorkflowFileName="Utility Files\InitAllSettings.xaml"
+ WorkflowFileName="Utility_Files\InitAllSettings.xaml"

- <InArgument x:Key="in_ConfigFile">[String.Format("Data\Config_{0}.xlsx",umgebung)]
+ <InArgument x:Key="in_ConfigFile">[String.Format("Data\Config_{0}.xlsx", umgebung)]
```

Der Pfad-Fix ist kein Kosmetikthema. Im ursprünglichen Template stand
`Utility Files\` mit Leerzeichen, der Ordner auf der Platte hiess aber
`Utility_Files`. Das knallt beim ersten Invoke.

### ReadSourceData.xaml

```diff
  <ui:ForEachFileX Filter="*.xlsx" ...>
+   <Assign DisplayName="Assign - Dokumentname setzen">
+     [out_dokumentname] = [CurrentFile.Name]
+   </Assign>
    <Assign DisplayName="Quellpfad aus Konfiguration"> ...
```

Zweiter Template-Fehler: `out_dokumentname` war als Output-Argument deklariert,
wurde aber nie zugewiesen. Das Feld `filename` im Queue-Item wäre also in jedem
Item leer geblieben. Der Generator ergänzt die Zuweisung, falls sie fehlt.

### PopulateQueue.xaml

Der grosse Umbau. Alles, was aus der Feldliste kommt, wird neu geschrieben:

| Stelle | Was passiert |
|---|---|
| `Sequence.Variables` | Wird komplett neu geschrieben: 13 fest im Skript hinterlegte Infrastruktur-Variablen (`queueName`, `primaerschluesselSpalte`, `datumsStempel`, `erstellZeitstempel`, `zeilenZaehler`, `neueSpalte`, `chunkGroesse`, `chunk_liste`, `fehlerZeilen_dt`, `anzErfolgreich`, `anzDuplikate`, `Reference`, `qItemInfo`) plus eine `String`-Variable pro aktivem Feld. Alle übrigen Variablen des Templates fallen weg. |
| `MultipleAssign` | Komplett neu, eine `AssignOperation` pro Feld, `Capacity` passend gesetzt. |
| `Assign QItemInfo` | Dictionary-Literal neu, in der Reihenfolge der Liste. |
| neu: `Assign - Reference erstellen` | wird vor dem QItemInfo-Assign eingehängt |
| neu: `Assign - Reference auf 128 Zeichen kuerzen` | nur wenn angehakt |
| `AddQueueItem` | `Reference`, `Priority`, `QueueType`, `FolderPath`, `ItemInformationCollection` |
| `Comment Out`-Block | wird entfernt |
| `Log: Start` | Meldung wird immer neu gesetzt und enthält den konfigurierten Queue-Schlüssel |
| übrige Logmeldungen | Meldungen, die alte Fach-Variablen erwähnen, werden auf `Reference` umgestellt |

Der `Comment Out`-Block im Original enthielt Ausdrücke mit `in_filename` und
`filenameReferenz` — beides Variablen, die es im Workflow gar nicht gab. Studio
validiert den Inhalt von auskommentierten Aktivitäten trotzdem und zeigt rote
Fehler an. Deshalb fliegt der Block raus.

### Data\Config_DEV.xlsx und Config_PRD.xlsx

Werden aus den Werten von Maske 4 komplett neu geschrieben, Sheet gemäss dem
Feld *Config-Sheet* auf Maske 1 (Vorgabe `Settings`), zwei Spalten
`Name` | `Value`. Alle Werte werden als Text geschrieben — eine Zahlenzelle
hätte aus `0815` die Zahl `815` gemacht, und UiPath liest die Werte ohnehin
über `.ToString()` aus dem `konfiguration_dict`.

---

## Die fünf Masken

### Maske 1 — Vorlage & Quelle

```
Template-Ordner     [C:\RPA\Templates\Dispatcher_Vorlage    ] [Durchsuchen]
Ziel-Ordner (neu)   [C:\RPA\Projekte\Zahlungen_Dispatcher   ] [Durchsuchen]
Projektname         [Zahlungen Dispatcher                   ]
Beschreibung        [Liest Zahlungsavise und befüllt die ... ]

Excel-Datei         [C:\RPA\Input\Zahlungen_Muster.xlsx      ] [Durchsuchen]
Sheet               [Übersicht                            ▼ ] [Spalten einlesen]

Config-Unterordner  [Data      ]
Config-Sheet        [Settings  ]
Default In_Umgebung [SZT2   ▼  ]
```

Wird die Excel-Datei über *Durchsuchen* gewählt, füllt sich die Sheet-Auswahl
von selbst. Ein von Hand eingetippter oder eingefügter Pfad tut das nicht — dann
bleibt die Liste leer und *Spalten einlesen* findet kein Sheet.

*Spalten einlesen* liest die erste Zeile des gewählten Sheets und legt daraus die
Feldliste an. Bei jedem Klick wird die Liste zurückgesetzt, samt aller auf
Maske 2 ergänzten Felder und der ganzen Reference auf Maske 3. Sind bereits
Felder oder Reference-Bausteine konfiguriert, fragt der Generator vorher nach.
Trotzdem: erst Excel und Sheet festzurren, dann einlesen, dann anpassen.

Beim Wählen des Template-Ordners liest der Generator im Hintergrund
`Data\Config_DEV.xlsx` und `Config_PRD.xlsx` und übernimmt deren Werte in
Maske 4. Fehlt eine der Dateien, macht das nichts.

Achtung bei der Reihenfolge: Sobald hier eine Quell-Excel gewählt wird — und bei
jedem Sheet-Wechsel im Dropdown — überschreibt der Generator auf Maske 4 die
Einträge `SourceDataPath` (Ordner der gewählten Datei) und `SourceSheetName`
(gewähltes Sheet), und zwar in der DEV- **und** der PRD-Spalte. Zuerst die Excel
wählen, dann Maske 4 anpassen — nicht umgekehrt.

Was mit den Spaltenüberschriften passiert:

```
Excel Zeile 1:   Vertrags-ID | Kunden Nr. | Kontounabhängige Zeile | Ausgangslage
                      │            │                │                    │
Queue-Schlüssel: Vertrags-ID | Kunden Nr. | Kontounabhängige Zeile | Ausgangslage
VB-Variable:      VertragsID  | KundenNr   | KontounabhaengigeZeile |  Ausgangslage
```

Der Schlüssel behält den Original-Namen, die Variable wird bereinigt: Umlaute
werden transliteriert (ä→ae, ö→oe, ü→ue, ß→ss), alles ausser `A-Z a-z 0-9 _`
fällt weg, Namen über 60 Zeichen werden auf 60 gekürzt, und doppelte Namen
bekommen ein Zählsuffix (`KundenNr`, `KundenNr2`) — angehängt nach dem Kürzen.
Der Zählsuffix greift auch gegen die 13 Infrastruktur-Variablen: eine Spalte
`Reference` wird zu `Reference2`, sonst gäbe es die Variable zweimal und Studio
öffnete den Workflow nicht mehr.

Spalten ohne Überschrift werden übersprungen und im Statustext gemeldet. Sie
bekommen bewusst keinen Ersatznamen: ein erfundenes `Spalte3` gibt es in der
DataTable nicht, und `CurrentRow("Spalte3")` liesse den Dispatcher beim ersten
Datensatz abbrechen.

### Maske 2 — Felder & Reihenfolge

Siehe [Queue-Item-Felder im Detail](#queue-item-felder-im-detail). Das ist der
Kern des Werkzeugs und hat einen eigenen Abschnitt.

### Maske 3 — Queue-Referenz

Siehe [Die Queue-Referenz](#die-queue-referenz).

### Maske 4 — Konfiguration DEV / PRD

Siehe [Konfiguration DEV / PRD](#konfiguration-dev--prd).

### Maske 5 — Bestätigung

Oben die vollständige Zusammenfassung inklusive dem fertigen
`qItemInfo`-Dictionary, so wie es gleich in die `.xaml` geschrieben wird. Unten
das Protokoll, das nach dem Erstellen jeden einzelnen Schritt auflistet. Die
Trennlinie dazwischen ist ziehbar.

Wer auf Maske 5 noch etwas ändert (Häkchen, Umschalten auf eine andere Maske und
zurück), bekommt die Zusammenfassung automatisch neu gebaut. Der Button
*Zusammenfassung aktualisieren* erzwingt es von Hand.

*Dispatcher erstellen* ist erst anklickbar, wenn das Häkchen darunter gesetzt
ist. Danach werden die Masken 1–4 noch einmal alle geprüft; die erste
unvollständige wird wieder geöffnet.

Ist der Zielordner nicht leer, kommt vorher eine Rückfrage. Sie ist ernst
gemeint: `Main.xaml`, `project.json` und die Utility-Workflows werden durch die
Template-Versionen ersetzt, Unterordner wie `Utility_Files` und `Data` werden
vorher **vollständig gelöscht**, `Config_DEV.xlsx` und `Config_PRD.xlsx` werden
neu geschrieben, und das Projekt bekommt eine neue `projectId`. Eigene
Änderungen in diesem Ordner sind danach weg — es wird nichts gesichert.

Der Zielordner darf nicht der Template-Ordner selbst sein und auch nicht darin
liegen; beides bricht mit einer Meldung ab, statt das Template zu überschreiben.

Bricht die Generierung mittendrin ab, bleibt der halbfertige Stand liegen — der
Generator legt dann eine Datei `GENERIERUNG_UNVOLLSTAENDIG.txt` in den
Zielordner. Liegt die dort, darf das Projekt nicht verwendet werden.

Meldet das Protokoll Zeilen mit `ACHTUNG`, verweist der generierte Workflow noch
auf einen Bezeichner aus dem Template, der nicht mehr deklariert ist. Der
Generator ersetzt so etwas bewusst nicht mehr automatisch — das hat früher stille
Falschzuordnungen erzeugt. Diese Stellen gehören in Studio angeschaut.

---

## Queue-Item-Felder im Detail

Das ist der Teil, der am meisten Zeit spart und am meisten schiefgehen kann.

### Was ein Queue-Item eigentlich ist

```
┌─ Queue Item ────────────────────────────────────────────────────┐
│                                                                 │
│  Reference   "V-88213_20260901143012"                           │
│              ↑ eine Zeile, max. 128 Zeichen, im Orchestrator     │
│                sichtbar und durchsuchbar                        │
│                                                                 │
│  Specific Data (= das qItemInfo-Dictionary)                     │
│  ┌───────────────────┬───────────────────────────────────────┐  │
│  │ filename          │ "Zahlungen_2026-09-01.xlsx"           │  │
│  │ CreateTS          │ 01.09.2026 14:30:12   (DateTime!)     │  │
│  │ VertragsID        │ "V-88213"                             │  │
│  │ KundenNr          │ "4711"                                │  │
│  │ Ausgangslage      │ "offen"                               │  │
│  │ Mandant           │ "SZKB"                                │  │
│  │ Lauf              │ "2026-09-01"                          │  │
│  └───────────────────┴───────────────────────────────────────┘  │
│              ↑ beliebig viele Schlüssel, das ist die Nutzlast    │
└─────────────────────────────────────────────────────────────────┘
```

Die Feldliste auf Maske 2 beschreibt genau den unteren Kasten. Eine Zeile in der
Liste = ein Schlüssel im Specific Data.

### Die Liste

```
┌──────┬──────────┬────────────────────────┬─────────────────────┬────────────────────┬──────────────┐
│Aktiv │ Typ      │ Excel-Spalte / Feld    │ Queue-Referenz      │ Wert / VB-Ausdruck │ Variable     │
├──────┼──────────┼────────────────────────┼─────────────────────┼────────────────────┼──────────────┤
│  ☑   │ Auto     │ (automatisch: Datei..) │ filename            │ in_dokumentname    │ @@FILENAME@@ │
│  ☑   │ Auto     │ (automatisch: Erstel.) │ CreateTS            │ DateTime.Now       │ @@CREATETS@@ │
│  ☑   │ Spalte   │ Vertrags-ID            │ VertragsID          │                    │ VertragsID   │
│  ☑   │ Spalte   │ Kunden Nr.             │ KundenNr            │                    │ KundenNr     │
│  ☐   │ Spalte   │ Interne Notiz          │ InterneNotiz        │                    │ InterneNotiz │
│  ☑   │ Spalte   │ Ausgangslage           │ Status              │                    │ Ausgangslage │
│  ☑   │ Text     │ (eigenes Feld)         │ Mandant             │ SZKB               │ Mandant      │
│  ☑   │ Ausdruck │ (eigenes Feld)         │ Lauf                │ DateTime.Now.To... │ Lauf         │
└──────┴──────────┴────────────────────────┴─────────────────────┴────────────────────┴──────────────┘
   ▲        ▲                ▲                       ▲                     ▲               ▲
   │        │                │                       │                     │               │
abwählbar  siehe          bei Spalten:          frei wählbar,        nur bei Text     wird berechnet,
           unten          muss exakt der        das ist der          und Ausdruck     nicht editierbar
                          Excel-Überschrift     Schlüssel im         editierbar
                          entsprechen           Queue-Item
```

Die rot unterstrichene Spalte *Queue-Referenz* ist immer editierbar. *Excel-Spalte*
und *Wert* ebenfalls, sofern der Feldtyp es zulässt. Änderungen greifen sofort beim
Tippen, man muss nichts bestätigen.

### Die vier Feldtypen

#### `Auto` — filename und CreateTS

Die beiden stehen fix an Position 1 und 2 und lassen sich weder entfernen noch
abwählen. Der Schlüsselname ist aber änderbar: wer im Performer lieber
`Dateiname` und `Erstellt` liest, schreibt das einfach in die Spalte
*Queue-Referenz*.

| | `filename` | `CreateTS` |
|---|---|---|
| Wert | `in_dokumentname` | `DateTime.Now` |
| Typ im Queue-Item | String | **DateTime** |
| Inhalt | Dateiname inkl. Endung, z. B. `Zahlungen_2026-09-01.xlsx` | Zeitpunkt, zu dem das Item gebaut wurde |

`CreateTS` ist der einzige Wert im ganzen Dictionary, der kein String ist. Das
ist Absicht: als echtes `DateTime` kann man im Performer damit rechnen, ohne
vorher zu parsen. Im Orchestrator-UI erscheint es trotzdem als Datumstext.

Warum die beiden überhaupt Pflicht sind: Wenn drei Wochen später jemand fragt,
warum ein Item so komisch aussieht, will man ohne Rätselraten wissen, aus welcher
Datei und aus welchem Lauf es stammt. Kostet zwei Schlüssel, spart eine
Stunde Suchen.

#### `Spalte` — der Normalfall

Kommt aus *Spalten einlesen*. In *Excel-Spalte / Feld* steht die
Original-Überschrift; sie muss zeichengenau stimmen, inklusive Leerzeichen,
Punkten und Umlauten. Wenn im Excel `Kunden-Nr.` steht, muss dort `Kunden-Nr.`
stehen — nicht `Kunden Nr`.

Erzeugt wird:

```xml
<ui:AssignOperation sap2010:WorkflowViewState.IdRef="AssignOperation_1">
  <ui:AssignOperation.To>
    <OutArgument x:TypeArguments="x:String">[KundenNr]</OutArgument>
  </ui:AssignOperation.To>
  <ui:AssignOperation.Value>
    <InArgument x:TypeArguments="x:String">[CurrentRow("Kunden-Nr.").ToString]</InArgument>
  </ui:AssignOperation.Value>
</ui:AssignOperation>
```

Leere Zellen ergeben `""`, nicht `Nothing` — `DBNull.Value.ToString()` liefert
einen Leerstring. Man kann also gefahrlos `.Trim` darauf anwenden. Existiert die
Spalte dagegen gar nicht, wirft `CurrentRow("...")` beim ersten Datensatz eine
`ArgumentException`. Das `MultipleAssign` liegt **ausserhalb** des Try/Catch —
im Template ist nur `AddQueueItem` davon umschlossen. Der ganze Dispatcher-Lauf
bricht damit ab, es wird keine einzige Zeile hochgeladen.

#### `Text` — fester Wert für alle Items

Über *+ Feld (fester Text)*. In der Spalte *Wert / VB-Ausdruck* steht reiner
Text, keine Anführungszeichen. Der Generator setzt die Quotes selbst und
verdoppelt vorhandene korrekt:

Eingabe `Bank "ZKB" & Co` wird zu

```xml
<ui:AssignOperation.Value>
  <InArgument x:TypeArguments="x:String">["Bank ""ZKB"" &amp; Co"]</InArgument>
</ui:AssignOperation.Value>
```

Typische Anwendungen: Mandant, Prozesskürzel, Schema-Version des Queue-Items,
ein fester Kostenstellen-Code.

#### `Ausdruck` — berechneter Wert

Über *+ Feld (VB-Ausdruck)*. Hier steht VB.NET-Code, der einmal pro Excel-Zeile
ausgewertet wird. Der Generator klammert ihn und hängt `.ToString` an:

Eingabe `DateTime.Now.ToString("yyyy-MM-dd")` wird zu

```xml
<ui:AssignOperation.Value>
  <InArgument x:TypeArguments="x:String">[(DateTime.Now.ToString("yyyy-MM-dd")).ToString]</InArgument>
</ui:AssignOperation.Value>
```

Verfügbar sind: alle Variablen aus der Feldliste (aber nur die, die **weiter
oben** stehen — dazu gleich mehr), `in_dokumentname`, `CurrentRow`,
`in_konfiguration_dict` und alles aus den importierten Namespaces
(`System`, `System.Linq`, `Microsoft.VisualBasic`, …).

### Die Reihenfolge ist nicht nur Kosmetik

Das `MultipleAssign` arbeitet die Liste von oben nach unten ab. Ein
`Ausdruck`-Feld, das auf eine andere Variable zugreift, funktioniert deshalb nur,
wenn diese Variable **darüber** steht.

```
   ✓ funktioniert                        ✗ ergibt Leerstring
   ┌────────────────────────┐            ┌────────────────────────┐
   │ 3. VertragsID  Spalte  │            │ 3. Voll        Ausdruck│  ← wird zuerst
   │ 4. KundenNr    Spalte  │            │    VertragsID & KundenNr│    ausgewertet
   │ 5. Voll        Ausdruck│            │ 4. VertragsID  Spalte  │
   │    VertragsID & KundenNr│           │ 5. KundenNr    Spalte  │
   └────────────────────────┘            └────────────────────────┘
```

Keine Fehlermeldung, kein Absturz — einfach ein leeres Feld. Die Variablen sind
zu dem Zeitpunkt schon deklariert, aber noch nicht gefüllt. Mit *Nach oben* /
*Nach unten* verschieben; die beiden Auto-Felder bleiben dabei immer oben.

Die Reihenfolge bestimmt ausserdem, wie das Dictionary im Code steht — also auch,
wie die Felder in der Orchestrator-Detailansicht sortiert erscheinen. Verwandte
Felder nebeneinander zu legen macht die Fehlersuche später deutlich angenehmer.

### Felder abwählen statt löschen

Das Häkchen in *Aktiv* nimmt ein Feld aus der Generierung — es verschwindet aus
den Variablen, aus dem `MultipleAssign` und aus dem Dictionary. Die Zeile bleibt
aber in der Liste stehen. Praktisch, wenn man beim Erstellen zwischen zwei
Varianten vergleicht.

*Feld entfernen* löscht die Zeile ganz. Mehrfachauswahl mit Strg oder Shift
funktioniert. Die beiden Auto-Felder ignorieren beides.

*Weiter* verlangt mindestens **ein** aktives Feld zusätzlich zu `filename` und
`CreateTS` — ein Queue-Item, das nur aus diesen beiden besteht, trägt keine
Fachdaten. Ausserdem müssen alle Queue-Schlüssel eindeutig sein.

### Was tatsächlich in der .xaml landet

Aus der Beispielliste oben (ohne *Interne Notiz*, weil abgewählt):

```vb
[New Dictionary(Of String, Object) From {
	{"filename", in_dokumentname},
	{"CreateTS", DateTime.Now},
	{"VertragsID", VertragsID},
	{"KundenNr", KundenNr},
	{"Status", Ausgangslage},
	{"Mandant", Mandant},
	{"Lauf", Lauf}
}]
```

Beachtenswert an der Zeile `{"Status", Ausgangslage}`: der Schlüssel im
Queue-Item heisst `Status`, die Excel-Spalte und die interne Variable heissen
`Ausgangslage`. Genau dafür ist die Spalte *Queue-Referenz* da — der Performer
muss die Excel-Struktur nicht kennen.

### Zugriff im Performer

```vb
' Strings
in_TransactionItem.SpecificContent("VertragsID").ToString
in_TransactionItem.SpecificContent("KundenNr").ToString.Trim

' CreateTS ist ein echtes DateTime
CDate(in_TransactionItem.SpecificContent("CreateTS"))
DateDiff(DateInterval.Hour, CDate(in_TransactionItem.SpecificContent("CreateTS")), Now)

' Defensiv, falls ein Schlüssel mal fehlt (z. B. Items aus einer älteren Version)
If in_TransactionItem.SpecificContent.ContainsKey("Mandant") Then ...
```

Der Schlüssel ist **case-sensitiv**. `SpecificContent("kundennr")` findet
`KundenNr` nicht.

### Kochbuch

Fertige Ausdrücke für `+ Feld (VB-Ausdruck)`:

```vb
' Datum und Zeit
DateTime.Now.ToString("yyyy-MM-dd")                    ' 2026-09-01
DateTime.Now.ToString("yyyyMMddHHmmss")                ' 20260901143012
DateTime.Now.ToString("yyyy-MM")                       ' Abrechnungsperiode
DateTime.Now.AddDays(14).ToString("dd.MM.yyyy")        ' Fälligkeit in 14 Tagen

' Aus dem Dateinamen
Path.GetFileNameWithoutExtension(in_dokumentname)      ' ohne .xlsx
in_dokumentname.Split("_"c)(1)                         ' zweiter Namensteil

' Aus der Konfiguration (steht in Config_DEV/PRD.xlsx)
in_konfiguration_dict("QueueName").ToString
in_konfiguration_dict("SourceSheetName").ToString

' Aus anderen Feldern (Feld muss weiter oben stehen!)
VertragsID.Trim & "-" & KundenNr.Trim
If(String.IsNullOrWhiteSpace(Ausgangslage), "unbekannt", Ausgangslage.Trim)
KundenNr.PadLeft(8, "0"c)                              ' 4711 -> 00004711

' Spalte direkt lesen, ohne sie als eigenes Feld zu führen
CurrentRow("Betrag").ToString.Replace(",", ".")
Decimal.Parse(CurrentRow("Betrag").ToString).ToString("F2")

' Sortier- und Gruppierungshilfen
Ausgangslage.Trim.ToLower                              ' normalisiert für Filter
Guid.NewGuid.ToString("N")                             ' technische Eindeutigkeit
Environment.UserName                                   ' wer hat den Lauf gestartet
```

### Stolpersteine

**Schlüssel doppelt.** Der Generator lässt sich nicht weiterklicken, wenn zwei
aktive Felder denselben Queue-Schlüssel haben — beim Dictionary-Aufbau würde das
zur Laufzeit eine `ArgumentException` geben. Gross-/Kleinschreibung wird bei
dieser Prüfung ignoriert, weil sich `KundenNr` und `kundennr` sonst nicht
nur beim Lesen beissen.

**Anführungszeichen in Ausdrücken.** Bei Typ `Ausdruck` gilt VB-Syntax, also
`"Text"` mit einfachen doppelten Anführungszeichen — der Generator fasst den
Ausdruck nicht an. Bei Typ `Text` dagegen nichts escapen, das erledigt der
Generator.

**Umlaute in Schlüsseln** sind erlaubt (`Kontounabhängige Zeile` funktioniert als
Dictionary-Key), in Variablennamen werden sie ersetzt. Das ist kein Widerspruch:
der Schlüssel ist ein String, der Variablenname VB-Code.

**Alles wird String.** Ausser `CreateTS`. Wer eine Zahl oder ein Datum typisiert
in der Queue braucht, formatiert im Ausdruck sauber (`"F2"`, `"yyyy-MM-dd"`) und
parst im Performer zurück. Das ist bewusst so — gemischte Typen im Specific Data
machen die Fehlersuche im Orchestrator-UI unangenehm.

**Sehr viele Felder.** Ab ungefähr 40 Feldern wird die
`Assign QItemInfo`-Aktivität in Studio unhandlich breit. Funktioniert weiterhin,
liest sich nur schlecht. Falls nur die Hälfte der Spalten im Performer wirklich
gebraucht wird, den Rest abwählen.

---

## Die Queue-Referenz

Die Reference ist die eine Zeile, die man im Orchestrator sieht, bevor man ein
Item aufklappt. Sie ist durchsuchbar und — falls die Queue auf *Enforce unique
references* steht — auch der Duplikatschutz. Mindestens ein Baustein ist Pflicht;
ohne einen kommt man nicht von Maske 3 weiter.

```
VERFÜGBARE BAUSTEINE              AUFBAU DER REFERENCE
┌────────────────────────┐        ┌────────────────────────────┐
│ [filename]             │        │ [VertragsID]  (VertragsID) │
│ [CreateTS]             │  Hin-  │ "_"                        │
│ [VertragsID] (Vertr..) │─zufü-─▶│ {Zeitstempel yyyyMMddHH..} │
│ [KundenNr]   (Kunde..) │  gen   │                            │
│ {Zeitstempel yyyyMM..} │        │  ▲ ▼ verschieben           │
│ {Datum yyyy-MM-dd}     │        └────────────────────────────┘
│ {Neue GUID}            │
│ {Trennzeichen _}       │        Eigener Text: [__________] [Text einfügen]
│ {Trennzeichen -}       │        ☑ Feldwerte trimmen
│ {Trennzeichen |}       │        ☑ auf 128 Zeichen kürzen
└────────────────────────┘

ERZEUGTER VB-AUSDRUCK
[VertragsID.Trim & "_" & DateTime.Now.ToString("yyyyMMddHHmmss")]
  danach:  [If(Reference.Length > 128, Reference.Substring(0, 128), Reference)]
```

Die Vorschau unten aktualisiert sich bei jeder Änderung. Was dort steht, landet
1:1 in der `.xaml`.

Nur **aktive** Felder erscheinen als Bausteine. Die linke Liste wird neu
aufgebaut bei *Spalten einlesen*, *Alle auswählen* / *Alle abwählen*,
*Feld entfernen* und beim Klick auf *Weiter* aus Maske 2. Die rechte Liste wird
bei *Spalten einlesen* geleert; bei *Alle abwählen* und *Feld entfernen* werden
gezielt die Bausteine entfernt, deren Feld nicht mehr aktiv ist.

Das *Aktiv*-Häkchen einzeln umzuschalten löst keinen dieser Wege aus — die
rechte Liste bleibt dann stehen. Spätestens beim Verlassen von Maske 3 wird das
aber geprüft: zeigt ein Baustein auf ein nicht mehr aktives Feld, kommt eine
Warnung statt einer `.xaml` mit einem Verweis auf eine Variable, die es im
generierten Workflow nicht mehr gibt.

*Feldwerte trimmen* hängt `.Trim` an jeden Feld-Baustein. Bei Excel-Exporten aus
SAP oder Avaloq ist das praktisch immer richtig.

*Auf 128 Zeichen kürzen* erzeugt einen zweiten Assign nach dem ersten. Ohne das
lehnt der Orchestrator zu lange Referenzen ab, und zwar erst beim Upload — also
mitten im Lauf, pro betroffener Zeile eine Exception.

### Was eine gute Reference ausmacht

```
schlecht    "Zeile 1", "Zeile 2", ...        sagt nichts, kollidiert zwischen Läufen
schlecht    Guid.NewGuid.ToString            eindeutig, aber nicht suchbar
brauchbar   VertragsID                       suchbar, aber Duplikate über Läufe
gut         VertragsID & "_" & yyyyMMdd      suchbar und pro Tag eindeutig
gut         Dateiname & "_" & Zeilenschlüssel wenn die Fachnummer fehlt
```

Wenn *Enforce unique references* auf der Queue aktiv ist und eine Reference
zweimal vorkommt, wirft `AddQueueItem` eine Exception. Der Generator baut den
Aufruf in ein `Try/Catch` (kommt aus dem Template), die Zeile wird also
protokolliert und der Lauf geht weiter. Man sieht es dann nur im Log — es lohnt
sich, nach dem ersten Lauf kurz reinzuschauen.

---

## Konfiguration DEV / PRD

```
┌──────────────────┬─────────────────────────────┬──────────────────────────┐
│ Name             │ Wert DEV (SZT1/SZT2)        │ Wert PRD (SZP)           │
├──────────────────┼─────────────────────────────┼──────────────────────────┤
│ SourceDataPath   │ C:\RPA\DEV\Input\           │ P:\RPA\PRD\Input\        │
│ SourceSheetName  │ Übersicht                   │ Übersicht                │
│ QueueFolder      │ arnenn@win's workspace      │ Shared                   │
│ QueueName        │ Zahlungen_DEV               │ Zahlungen                │
│ ChunkSize        │ 10000                       │ 10000                    │
└──────────────────┴─────────────────────────────┴──────────────────────────┘
```

*Weiter* verlangt hier, dass `SourceDataPath`, `SourceSheetName` und `QueueName`
für **DEV** ausgefüllt sind. `QueueFolder` wird nicht geprüft, und die
PRD-Spalte gar nicht — die kann man auch später in Studio nachtragen.

Die vier obersten Zeilen sind fix und können nicht umbenannt oder gelöscht
werden. `ReadSourceData` greift namentlich auf `SourceDataPath` und
`SourceSheetName` zu, `PopulateQueue` auf `QueueName` und `QueueFolder`.
`InitAllSettings` liest keinen dieser Schlüssel einzeln — es füllt nur das
`konfiguration_dict`.
Eigene Zeilen kommen über *+ Zeile hinzufügen* dazu und stehen im Workflow als
`in_konfiguration_dict("MeinSchluessel")` zur Verfügung.

`SourceDataPath` gehört mit abschliessendem Backslash geschrieben und zeigt auf
einen **Ordner**, nicht auf eine Datei. Der Dispatcher nimmt sich daraus alle
Dateien, die auf den Dateifilter passen.

Darunter auf derselben Maske:

- **Dateifilter Input-Ordner** — landet in `ForEachFileX`. Standard `*.xlsx`.
  Wer nur bestimmte Dateien will: `Zahlungen_*.xlsx`.
- **Queue-Item Priorität** — `Low`, `Normal`, `High`. Wird fest in
  `AddQueueItem` geschrieben, nicht über die Config.
- **Schlüssel Queue-Name / Queue-Folder** — falls die Config-Zeilen im eigenen
  Template anders heissen als `QueueName` und `QueueFolder`.

*Werte aus Template übernehmen* liest die beiden Config-Dateien des Templates
nochmal ein und überschreibt die Tabelle. Nützlich, wenn man zwischendurch
etwas kaputtgeklickt hat.

---

## Nach dem Generieren: Checkliste in Studio

Der Generator erzeugt gültige Workflows, aber er ersetzt keinen Blick drauf.

1. Projekt in Studio öffnen. Es dürfen keine roten Fehler stehen. Kommen welche,
   liegt es fast immer an einem `Ausdruck`-Feld mit Tippfehler.
2. `PopulateQueue.xaml` öffnen, `Assign QItemInfo` aufklappen und gegen die
   Zusammenfassung von Maske 5 halten.
3. `Assign - Reference erstellen` prüfen — steht der erwartete Ausdruck drin?
4. `Data\Config_DEV.xlsx` öffnen. Zeigt `SourceDataPath` auf einen Ordner, der
   auf dem Robot-Rechner tatsächlich existiert und erreichbar ist?
5. Queue im Orchestrator anlegen, falls noch nicht vorhanden. Der Generator
   legt keine Queues an.
6. Einmal mit `In_Umgebung = SZT2` und **einer** Testdatei laufen lassen.
7. Erstes Item im Orchestrator aufklappen: Reference plausibel? Alle erwarteten
   Schlüssel da? `filename` gefüllt?

Punkt 7 ist der eigentliche Test. Wenn das Item stimmt, stimmen alle.

---

## Grenzen und bekannte Eigenheiten

**Mehrere Dateien im Input-Ordner.** `ReadSourceData` läuft über alle Dateien,
die auf den Filter passen, überschreibt `out_quelldaten_dt` aber bei jeder
Iteration. `PopulateQueue` wird danach genau einmal aufgerufen. Ergebnis: nur die
**letzte** Datei landet in der Queue. Das ist Verhalten des Templates, nicht des
Generators — er lässt die Logik bewusst unangetastet. Wer wirklich mehrere
Dateien pro Lauf braucht, muss `Main.xaml` umbauen und den `Populate Queue`-State
in die Dateischleife ziehen.

**Nur OpenXML beim Einlesen.** Sheets und Spaltenüberschriften kommen aus
`.xlsx` und `.xlsm`. `.xls` (altes Binärformat) und `.xlsb` kann der Generator nicht öffnen.
Der generierte Workflow selbst kann via `ReadRange` durchaus `.xls` lesen — man
muss die Spalten dann nur von Hand eintragen statt sie einlesen zu lassen.

**Kopfzeile muss Zeile 1 sein.** Wer im Excel zwei Titelzeilen über der Tabelle
hat, bekommt diese als Spaltennamen. Entweder vorher aufräumen oder die
Feldliste manuell korrigieren.

**Config-Dateien werden neu geschrieben,** nicht bearbeitet. Formatierung,
Kommentarspalten und zusätzliche Sheets in den alten `Config_*.xlsx` gehen dabei
verloren. Sie enthalten normalerweise nur Name/Value, deshalb ist das
verschmerzbar — aber es ist gut zu wissen.

**Nicht implementiert, obwohl die Variablen da sind.** `chunkGroesse`,
`chunk_liste`, `anzDuplikate`, `anzErfolgreich`, `fehlerZeilen_dt`,
`primaerschluesselSpalte`, `queueName`, `datumsStempel`, `erstellZeitstempel`,
`zeilenZaehler` (Default `2`) und `neueSpalte` (`DataColumn`) existieren als
Variablen in `PopulateQueue.xaml`, es gibt aber keine Aktivitäten dazu. Der Generator lässt sie stehen, damit man
darauf aufbauen kann, ohne sie neu anzulegen. Ebenso werden `ChunkSize` und
`PrimaryKeyColumn` aus der Config zwar geladen, aber nirgends ausgewertet.

**Keine Orchestrator-Anbindung.** Queues, Assets und Ordner muss man selbst
anlegen. Der Generator schreibt nur die Namen in die Config.

---

## Fehlersuche

**„PopulateQueue.xaml wurde im kopierten Projekt nicht gefunden."**
Der Template-Ordner ist nicht der Projektordner, sondern eine Ebene darüber oder
darunter. Es muss der Ordner sein, in dem `project.json` liegt.

**„Sheet 'X' wurde in '...' nicht gefunden."**
Beim Lesen der Config-Dateien: Der Config-Sheet-Name auf Maske 1 (Standard
`Settings`) passt nicht zum tatsächlichen Sheet in `Config_DEV.xlsx`.

**„In der ersten Zeile des Sheets wurden keine Spaltenüberschriften gefunden."**
Falsches Sheet gewählt, oder die Tabelle beginnt nicht in Zeile 1.

**Die Zusammenfassung auf Maske 5 zeigt einen Fehler statt der Übersicht.**
Der komplette Fehlertext samt Stacktrace steht dann im Feld. In aller Regel ist
ein `Ausdruck`-Feld schuld.

**Studio zeigt nach dem Generieren rote Fehler in `PopulateQueue.xaml`.**
Fast immer ein `Ausdruck`-Feld: Tippfehler, fehlende Klammer, oder Zugriff auf
eine Variable, die in der Liste weiter unten steht. Der Generator prüft
VB-Ausdrücke nicht — das kann nur der Studio-Compiler.

**Alle Items haben ein leeres `filename`.**
Der Generator ergänzt die Zuweisung nur, wenn sie fehlt. Wurde `out_dokumentname`
im Template zwischenzeitlich anders zugewiesen (z. B. auf eine Variable, die nie
gefüllt wird), bleibt das so. In `ReadSourceData.xaml` nachsehen.

**„The reference is too long" beim Upload.**
Häkchen *auf 128 Zeichen kürzen* auf Maske 3 war nicht gesetzt. Entweder neu
generieren oder in `PopulateQueue.xaml` von Hand einen zweiten Assign ergänzen.

**Das Fenster geht gar nicht auf.**
Meist die Execution Policy. Mit `-ExecutionPolicy Bypass` starten. Falls die
Konsole kommentarlos zurückkommt: das Skript hat sich in einen STA-Prozess neu
gestartet, das Fenster erscheint dann kurz danach.

**Die Oberfläche sieht nach Standard-Windows aus statt nach Hausschrift.**
*Futura Now Text* ist auf dem Rechner nicht installiert. Rein optisch, ohne
Auswirkung auf die Funktion.

---

## Wie es intern funktioniert

Für den Fall, dass jemand das Skript anpassen muss.

### Aufbau

```
New-UiPathDispatcher.ps1
 0  STA-Prüfung, Neustart mit -STA falls nötig
 1  Datenmodell           C# via Add-Type: FieldItem, ConfigItem, RefBlock
                          (INotifyPropertyChanged, damit die ListViews live binden)
 2  Hilfsfunktionen       Umlaut-Transliteration, VB-Identifier, XML-/VB-Escaping
 3  Excel lesen           Get-XlsxSharedString, Get-XlsxCellValue,
                          Get-XlsxSheet, Get-XlsxHeader, Get-XlsxNameValue
 4  Excel schreiben       New-ConfigXlsx
 5  XAML-Werkzeug         Open-XamlDocument, Namespace-Manager, Save-XamlDocument
 6  Generierung           Update-ProjectJson, Update-MainXaml,
                          Update-ReadSourceDataXaml, Update-PopulateQueueXaml,
                          Invoke-DispatcherGeneration
 7  Oberfläche            WPF-XAML als Here-String, Controls, Event-Handler
```

Abschnitte 1–6 haben keine Abhängigkeit zur Oberfläche. Man kann sie einzeln
dot-sourcen und `Invoke-DispatcherGeneration` mit einer Hashtable aufrufen, wenn
man den Generator irgendwann headless in eine Pipeline hängen will.

### Excel ohne Excel

`Get-XlsxSheet` öffnet die `.xlsx` als ZIP, liest `xl/workbook.xml` für die
Sheet-Namen und löst die `r:id` über `xl/_rels/workbook.xml.rels` zum
tatsächlichen Sheet-Pfad auf. `Get-XlsxHeader` lädt danach `xl/sharedStrings.xml`
und streamt das Sheet mit einem `XmlReader`, bricht nach der ersten `<row>` ab
und übersetzt Zellreferenzen (`C1` → Index 2) auf Spaltenpositionen. Vom Sheet
selbst wird also nie mehr als die erste Zeile geparst, egal wie gross die Datei
ist. Nur die Shared-String-Tabelle geht einmal komplett durch — die ist bei
breiten, aber normal langen Tabellen unkritisch.

`New-ConfigXlsx` baut eine minimale, gültige Arbeitsmappe von Grund auf:
`[Content_Types].xml`, `_rels/.rels`, `xl/workbook.xml`,
`xl/_rels/workbook.xml.rels`, `xl/styles.xml` und ein Sheet mit Inline-Strings
(`t="inlineStr"`). Kein `sharedStrings`-Zoo, keine Alt-Formatierung, die
mitgeschleppt werden müsste.

### XAML-Bearbeitung

Geladen wird mit `PreserveWhitespace = $true`, gespeichert über einen
`XmlWriter` mit `OmitXmlDeclaration = $true` und `Indent = $false`. Damit bleibt
die Formatierung des Originals erhalten und es kommt keine
`<?xml ...?>`-Zeile dazu — UiPath-`.xaml`-Dateien haben keine.

Neue Knoten werden gebaut, indem ein leeres Element erst in den Baum gehängt und
danach sein `InnerXml` gesetzt wird. Nur so lösen sich die Präfixe `ui:`, `x:`
und `sap2010:` gegen die Namespace-Deklarationen der Wurzel auf.

Zum Schluss läuft ein Prüfdurchgang: Bezeichner, die im Template als
Fach-Variablen existierten und im neuen Dispatcher nicht mehr deklariert werden,
aber noch in einem `InArgument`, `OutArgument` oder einer Logmeldung auftauchen,
werden als `ACHTUNG`-Zeile ins Protokoll geschrieben.

Ersetzt wird dabei **nichts**. Früher wurden diese Bezeichner im gesamten
Dokument per Regex durch `Reference` ersetzt, und das ging still schief: der
`QueueKey` und der `ColumnName` stehen ungefaltet im XML, ihr `VarName` dagegen
umlautgefaltet (`ae` statt `ä`). Eine Spalte `Kontounabhängige Zeile` galt damit
als alter Bezeichner, und das gerade erst erzeugte
`[CurrentRow("Kontounabhängige Zeile").ToString]` wurde zu
`[CurrentRow("Reference").ToString]` — zur Laufzeit *Column does not belong to
table*, bei zwei solchen Spalten zusätzlich ein doppelter Dictionary-Schlüssel.
Eine gemeldete Stelle in Studio zu korrigieren ist harmloser als eine still
falsch umgebogene.

### Erweitern

Ein neuer Feldtyp (sagen wir `Konfiguration`, der direkt aus dem
`konfiguration_dict` liest) braucht drei Eingriffe:

1. `Update-PopulateQueueXaml`, `switch ($f.Kind)` — den Ausdruck erzeugen
2. `New-QItemDictionary` — falls der Wert nicht über eine Variable laufen soll
3. Maske 2 — einen Button analog zu `btnFieldAddExp`

Eine zusätzliche Pflichtzeile in der Config wird in `Reset-ConfigRow` im Array
`$defaults` ergänzt, mit `L = $true` für nicht löschbar.
# Dispatcher-Generator
# Dispatcher-Generator
