<#
.SYNOPSIS
    UiPath Dispatcher Generator - WPF/XAML-Assistent, der ein bestehendes
    Dispatcher-Template kopiert und vollstaendig auf ein neues Quell-Excel
    anpasst (Spalten -> Queue-Item-Referenzen, Reference-Aufbau, Config DEV/PRD).

.DESCRIPTION
    Analysiert wurde folgendes Template-Geruest:

      Main.xaml                        StateMachine "Dispatcher Framework"
                                       Init -> Get Source Data -> Populate Queue -> End
                                       Argument In_Umgebung (SZT1/SZT2 -> DEV, SZP -> PRD)
      Utility_Files/InitAllSettings    liest Data\Config_{DEV|PRD}.xlsx, Sheet "Settings"
                                       Name/Value -> konfiguration_dict
      Utility_Files/ReadSourceData     ForEachFileX ueber konfiguration_dict("SourceDataPath")
                                       ReadRange Sheet = konfiguration_dict("SourceSheetName")
                                       -> out_quelldaten_dt / out_keineDaten / out_dokumentname
      Utility_Files/PopulateQueue      ForEachRow -> MultipleAssign (Spalte -> Variable)
                                       -> Assign qItemInfo (Dictionary)
                                       -> AddQueueItem (Reference, QueueType, FolderPath)

    Angepasst werden durch diesen Generator:
      * project.json          Name, Beschreibung, neue GUIDs
      * Main.xaml             Default In_Umgebung, korrekte Pfade der Utility-Files,
                              Config-Ordner / Config-Sheet
      * ReadSourceData.xaml   Dateifilter, fehlende Zuweisung von out_dokumentname
      * PopulateQueue.xaml    Variablen, MultipleAssign, qItemInfo-Dictionary,
                              Reference-Aufbau, Prioritaet, Logmeldungen
      * Data/Config_DEV.xlsx  SourceDataPath, SourceSheetName, QueueFolder, QueueName, ...
      * Data/Config_PRD.xlsx

.NOTES
    PowerShell 5.1+ (Windows), STA. Keine externen Module, kein Excel noetig -
    Excel-Dateien werden direkt ueber OpenXML (ZIP + XML) gelesen und geschrieben.
#>

[CmdletBinding()]
param(
    [string]$TemplatePath,
    [string]$OutputPath,
    # Wird vom STA-Neustart gesetzt und verhindert eine Neustart-Schleife.
    [switch]$NoRelaunch
)

# ===========================================================================
#  0. STA sicherstellen (WPF benoetigt Single Threaded Apartment)
# ===========================================================================
$ErrorActionPreference = 'Stop'

if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
    # Frueher scheiterte dieser Neustart lautlos: Start-Process ohne -PassThru, und
    # $ErrorActionPreference stand erst danach auf 'Stop'. Man sah ein kurz
    # aufblitzendes Fenster und sonst nichts.
    if ($NoRelaunch) {
        throw 'Dieses Skript benoetigt einen STA-Thread. Bitte mit "powershell.exe -STA -File New-UiPathDispatcher.ps1" starten.'
    }
    if ($PSVersionTable.PSEdition -eq 'Core') {
        throw 'Bitte Windows PowerShell 5.1 (powershell.exe) verwenden - PowerShell 7 kennt den Schalter -STA nicht.'
    }
    if ([string]::IsNullOrWhiteSpace($PSCommandPath)) {
        throw 'Der Skriptpfad ist nicht bekannt (Aufruf ueber die Pipeline?). Bitte die .ps1 direkt starten.'
    }
    # Bewusst nicht (Get-Process -Id $PID).Path: in der ISE oder einem eingebetteten
    # Host ist das ein Programm, das -File und -STA gar nicht kennt.
    $argList = @('-STA', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"", '-NoRelaunch')
    if ($TemplatePath) { $argList += @('-TemplatePath', "`"$TemplatePath`"") }
    if ($OutputPath)   { $argList += @('-OutputPath',   "`"$OutputPath`"") }
    $proc = Start-Process -FilePath 'powershell.exe' -ArgumentList $argList -PassThru
    if ($proc.WaitForExit(2000) -and $proc.ExitCode -ne 0) {
        throw "Der STA-Neustart ist fehlgeschlagen (Exit-Code $($proc.ExitCode))."
    }
    return
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# ===========================================================================
#  1. Datenmodell (INotifyPropertyChanged fuer Live-Binding im ListView)
# ===========================================================================
if (-not ('UiPathGen.FieldItem' -as [type])) {
Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;

namespace UiPathGen
{
    public class NotifyBase : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnChanged(string name)
        {
            var h = PropertyChanged;
            if (h != null) h(this, new PropertyChangedEventArgs(name));
        }
    }

    /// <summary>Eine Spalte oder ein manuell hinzugefuegtes Feld des Queue-Items.</summary>
    public class FieldItem : NotifyBase
    {
        private bool _selected = true;
        private string _kind = "Spalte";      // Spalte | Text | Ausdruck | Auto
        private string _columnName = "";
        private string _queueKey = "";
        private string _varName = "";
        private string _expression = "";
        private bool _locked = false;

        public bool Selected
        {
            get { return _selected; }
            set { if (_locked) { _selected = true; } else { _selected = value; } OnChanged("Selected"); }
        }
        public string Kind
        {
            get { return _kind; }
            set { _kind = value; OnChanged("Kind"); OnChanged("ValueEditable"); }
        }
        public string ColumnName
        {
            get { return _columnName; }
            set { _columnName = value; OnChanged("ColumnName"); }
        }
        /// <summary>Schluessel, unter dem der Wert im Queue-Item (Specific Data) landet.</summary>
        public string QueueKey
        {
            get { return _queueKey; }
            set { _queueKey = value; OnChanged("QueueKey"); }
        }
        /// <summary>VB-Variablenname im Workflow.</summary>
        public string VarName
        {
            get { return _varName; }
            set { _varName = value; OnChanged("VarName"); }
        }
        /// <summary>Literal (Kind=Text) bzw. VB-Ausdruck (Kind=Ausdruck).</summary>
        public string Expression
        {
            get { return _expression; }
            set { _expression = value; OnChanged("Expression"); }
        }
        /// <summary>Pflichtfeld (filename / CreateTS) - nicht loesch- und nicht abwaehlbar.</summary>
        public bool IsLocked
        {
            get { return _locked; }
            set { _locked = value; OnChanged("IsLocked"); OnChanged("IsUnlocked"); OnChanged("ValueEditable"); }
        }
        public bool IsUnlocked { get { return !_locked; } }
        public bool ValueEditable { get { return !_locked && _kind != "Spalte"; } }
    }

    /// <summary>Eine Zeile des Settings-Sheets in Config_DEV.xlsx / Config_PRD.xlsx.</summary>
    public class ConfigItem : NotifyBase
    {
        private string _name = "";
        private string _dev = "";
        private string _prd = "";
        private bool _locked = false;

        public string Name { get { return _name; } set { _name = value; OnChanged("Name"); } }
        public string DevValue { get { return _dev; } set { _dev = value; OnChanged("DevValue"); } }
        public string PrdValue { get { return _prd; } set { _prd = value; OnChanged("PrdValue"); } }
        public bool IsLocked { get { return _locked; } set { _locked = value; OnChanged("IsLocked"); } }
    }

    /// <summary>Ein Baustein der Queue-Item-Reference.</summary>
    public class RefBlock : NotifyBase
    {
        private string _display = "";
        private string _expression = "";
        private string _kind = "";     // Feld | Text | Sonder

        public string Display { get { return _display; } set { _display = value; OnChanged("Display"); } }
        public string Expression { get { return _expression; } set { _expression = value; OnChanged("Expression"); } }
        public string Kind { get { return _kind; } set { _kind = value; OnChanged("Kind"); } }
    }
}
'@
}

# ===========================================================================
#  2. Hilfsfunktionen - allgemein
# ===========================================================================

function Convert-Umlaut {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    # Bewusst als Paar-Liste (nicht Hashtable): PowerShell-Hashtable-Keys sind
    # case-insensitiv, 'ä' und 'Ä' wuerden kollidieren. String.Replace ist case-sensitiv.
    $t = $Text
    $map = @(
        @('ä', 'ae'), @('ö', 'oe'), @('ü', 'ue'),
        @('Ä', 'Ae'), @('Ö', 'Oe'), @('Ü', 'Ue'), @('ß', 'ss'),
        @('à', 'a'), @('á', 'a'), @('â', 'a'), @('è', 'e'), @('é', 'e'), @('ê', 'e'),
        @('ì', 'i'), @('í', 'i'), @('ò', 'o'), @('ó', 'o'), @('ù', 'u'), @('ú', 'u'),
        @('ç', 'c'), @('À', 'A'), @('É', 'E'), @('È', 'E'), @('Ç', 'C')
    )
    foreach ($p in $map) { $t = $t.Replace($p[0], $p[1]) }
    return $t
}

function ConvertTo-VbIdentifier {
    <# Erzeugt aus einem beliebigen Spaltennamen einen gueltigen VB-Variablennamen. #>
    param([string]$Name, [System.Collections.Generic.HashSet[string]]$Used)

    $s = Convert-Umlaut $Name
    $s = ($s -replace '[^A-Za-z0-9_]', '')
    if ([string]::IsNullOrWhiteSpace($s)) { $s = 'Feld' }
    if ($s -match '^[0-9]') { $s = 'f' + $s }
    if ($s.Length -gt 60) { $s = $s.Substring(0, 60) }

    if ($Used) {
        $base = $s; $i = 2
        while ($Used.Contains($s.ToLowerInvariant())) { $s = "$base$i"; $i++ }
        [void]$Used.Add($s.ToLowerInvariant())
    }
    return $s
}

function ConvertTo-XmlText {
    <#
      Escaping fuer XML-Textknoten UND fuer per String-Interpolation gebaute
      Attributwerte (New-ConfigXlsx baut workbook.xml als Text zusammen).
      Deshalb wird das Anfuehrungszeichen ebenfalls escaped - &quot; ist auch
      in einem Textknoten gueltig.
    #>
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return $Text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;')
}

function ConvertTo-VbLiteral {
    <#
      Erzeugt ein VB-String-Literal inkl. korrektem Quote-Escaping.
      Zeilenumbrueche werden zu Leerzeichen: VB.NET kennt keine mehrzeiligen
      Literale, ein CR/LF im Text ergaebe sonst einen Compilerfehler in Studio.
    #>
    param([string]$Text)
    if ($null -eq $Text) { $Text = '' }
    $t = $Text -replace '[\r\n\t]+', ' '
    return '"' + $t.Replace('"', '""') + '"'
}

function Get-ColumnIndexFromRef {
    param([string]$CellRef)
    $letters = ($CellRef -replace '[^A-Za-z]', '').ToUpperInvariant()
    if ([string]::IsNullOrEmpty($letters)) { return -1 }
    $idx = 0
    foreach ($ch in $letters.ToCharArray()) { $idx = $idx * 26 + ([int][char]$ch - 64) }
    return $idx - 1
}

# ===========================================================================
#  3. Excel lesen (OpenXML, ohne Excel-Installation)
# ===========================================================================

function Get-XlsxSharedString {
    <# Liest xl/sharedStrings.xml (falls vorhanden) als Liste ein. #>
    param(
        [Parameter(Mandatory)][System.IO.Compression.ZipArchive]$Zip,
        [Parameter(Mandatory)][string]$MainNs
    )

    $shared = New-Object System.Collections.ArrayList
    $ssEntry = $Zip.Entries | Where-Object { $_.FullName -eq 'xl/sharedStrings.xml' } | Select-Object -First 1
    if ($ssEntry) {
        $ssDoc = New-Object System.Xml.XmlDocument
        $s = $ssEntry.Open()
        try { $ssDoc.Load($s) } finally { $s.Dispose() }
        $sns = New-Object System.Xml.XmlNamespaceManager($ssDoc.NameTable)
        $sns.AddNamespace('m', $MainNs)
        foreach ($si in $ssDoc.SelectNodes('//m:si', $sns)) {
            $sb = New-Object System.Text.StringBuilder
            foreach ($t in $si.SelectNodes('.//m:t', $sns)) { [void]$sb.Append($t.InnerText) }
            [void]$shared.Add($sb.ToString())
        }
    }
    # Komma: sonst entrollt PowerShell die Liste zu Einzelwerten.
    return ,$shared
}

function Get-XlsxCellValue {
    <#
      Wert einer <c>-Zelle. t="s" verweist in die SharedStrings, t="inlineStr"
      traegt den Text im Element selbst, alles andere (Zahl, Datum, aber auch
      ein Formelergebnis t="str") steht direkt in <v>.
    #>
    param(
        [Parameter(Mandatory)]$Cell,
        [Parameter(Mandatory)]$Shared,
        [Parameter(Mandatory)][System.Xml.XmlNamespaceManager]$Nsm
    )

    $type = $Cell.GetAttribute('t')
    if ($type -eq 's') {
        $vNode = $Cell.SelectSingleNode('m:v', $Nsm)
        if ($vNode) {
            $si = 0
            if ([int]::TryParse($vNode.InnerText, [ref]$si) -and $si -ge 0 -and $si -lt $Shared.Count) {
                return $Shared[$si]
            }
        }
        return ''
    }
    if ($type -eq 'inlineStr') {
        $sb = New-Object System.Text.StringBuilder
        foreach ($t in $Cell.SelectNodes('.//m:t', $Nsm)) { [void]$sb.Append($t.InnerText) }
        return $sb.ToString()
    }
    $vNode = $Cell.SelectSingleNode('m:v', $Nsm)
    if ($vNode) { return $vNode.InnerText }
    return ''
}

function Get-XlsxSheet {
    <# Liefert die Sheets einer .xlsx als Objekte mit Name + Zip-Pfad. #>
    param([Parameter(Mandatory)][string]$Path)

    $result = New-Object System.Collections.ArrayList
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $mainNs = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        $relNs  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        $pkgRel = 'http://schemas.openxmlformats.org/package/2006/relationships'

        $wbEntry  = $zip.Entries | Where-Object { $_.FullName -eq 'xl/workbook.xml' } | Select-Object -First 1
        $relEntry = $zip.Entries | Where-Object { $_.FullName -eq 'xl/_rels/workbook.xml.rels' } | Select-Object -First 1
        if (-not $wbEntry) { throw "Ungueltige Excel-Datei (xl/workbook.xml fehlt): $Path" }

        $wbDoc = New-Object System.Xml.XmlDocument
        $s = $wbEntry.Open(); $wbDoc.Load($s); $s.Dispose()

        $relMap = @{}
        if ($relEntry) {
            $relDoc = New-Object System.Xml.XmlDocument
            $s = $relEntry.Open(); $relDoc.Load($s); $s.Dispose()
            $rns = New-Object System.Xml.XmlNamespaceManager($relDoc.NameTable)
            $rns.AddNamespace('p', $pkgRel)
            foreach ($r in $relDoc.SelectNodes('//p:Relationship', $rns)) {
                $relMap[$r.GetAttribute('Id')] = $r.GetAttribute('Target')
            }
        }

        $wns = New-Object System.Xml.XmlNamespaceManager($wbDoc.NameTable)
        $wns.AddNamespace('m', $mainNs)
        $wns.AddNamespace('r', $relNs)

        $i = 0
        foreach ($sh in $wbDoc.SelectNodes('//m:sheets/m:sheet', $wns)) {
            $i++
            $rid    = $sh.GetAttribute('id', $relNs)
            $target = $null
            if ($rid -and $relMap.ContainsKey($rid)) { $target = $relMap[$rid] }
            if (-not $target) { $target = "worksheets/sheet$i.xml" }
            $target = $target -replace '^/xl/', '' -replace '^/', ''
            if ($target -notlike 'xl/*') { $target = 'xl/' + $target }
            [void]$result.Add([pscustomobject]@{
                Name  = $sh.GetAttribute('name')
                Path  = $target
                Index = $i
            })
        }
    }
    finally { $zip.Dispose() }
    return $result
}

function Get-XlsxHeader {
    <# Liest die erste Zeile (Kopfzeile) eines Sheets und gibt die Spaltennamen zurueck. #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$SheetName
    )

    $mainNs = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $sheets = Get-XlsxSheet -Path $Path
    $sheet  = $sheets | Where-Object { $_.Name -eq $SheetName } | Select-Object -First 1
    if (-not $sheet) { throw "Sheet '$SheetName' wurde in '$([System.IO.Path]::GetFileName($Path))' nicht gefunden." }

    $headers = New-Object System.Collections.ArrayList
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        # --- Shared Strings (falls vorhanden) einlesen ---
        $shared = Get-XlsxSharedString -Zip $zip -MainNs $mainNs

        # --- Erste Zeile des Sheets streamen ---
        $shEntry = $zip.Entries | Where-Object { $_.FullName -eq $sheet.Path } | Select-Object -First 1
        if (-not $shEntry) { throw "Sheet-Daten '$($sheet.Path)' nicht gefunden." }

        $stream = $shEntry.Open()
        try {
            $reader = [System.Xml.XmlReader]::Create($stream)
            try {
                while ($reader.Read()) {
                    if ($reader.NodeType -eq [System.Xml.XmlNodeType]::Element -and $reader.LocalName -eq 'row') {
                        $sub = $reader.ReadSubtree()
                        $rowDoc = New-Object System.Xml.XmlDocument
                        $rowDoc.Load($sub)
                        $sub.Dispose()

                        $rns = New-Object System.Xml.XmlNamespaceManager($rowDoc.NameTable)
                        $rns.AddNamespace('m', $mainNs)

                        $cells = @{}
                        $maxIdx = -1
                        $auto = 0
                        foreach ($c in $rowDoc.SelectNodes('//m:c', $rns)) {
                            $ref = $c.GetAttribute('r')
                            $idx = Get-ColumnIndexFromRef $ref
                            if ($idx -lt 0) { $idx = $auto }
                            $auto = $idx + 1
                            $val = Get-XlsxCellValue -Cell $c -Shared $shared -Nsm $rns
                            $cells[$idx] = $val
                            if ($idx -gt $maxIdx) { $maxIdx = $idx }
                        }

                        for ($i = 0; $i -le $maxIdx; $i++) {
                            $v = ''
                            if ($cells.ContainsKey($i)) { $v = $cells[$i] }
                            # Leere Kopfzellen behalten den Leerstring. Frueher stand hier ein
                            # erfundenes "SpalteN" - das erzeugte eine Spalte, die es in der
                            # DataTable nicht gibt, und der Dispatcher brach zur Laufzeit ab.
                            # Interner Zeilenumbruch (Alt+Enter) wuerde spaeter das VB-Literal
                            # sprengen, deshalb hier schon zu einem Leerzeichen normalisieren.
                            $v = ($v -replace '\s+', ' ').Trim()
                            [void]$headers.Add($v)
                        }
                        break
                    }
                }
            }
            finally { $reader.Dispose() }
        }
        finally { $stream.Dispose() }
    }
    finally { $zip.Dispose() }

    # Nachlaufende Leerspalten entfernen. Bewusst nur exakt leere Eintraege -
    # ein Muster wie '^Spalte\d+$' wuerde eine echte Kopfzeile 'Spalte1' loeschen.
    while ($headers.Count -gt 0 -and [string]::IsNullOrEmpty($headers[$headers.Count - 1])) {
        $headers.RemoveAt($headers.Count - 1)
    }
    return $headers
}

function Get-XlsxNameValue {
    <# Liest ein Name/Value-Sheet (Settings) als Hashtable. #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$SheetName = 'Settings'
    )

    $mainNs = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $result = [ordered]@{}
    $sheets = Get-XlsxSheet -Path $Path
    $sheet  = $sheets | Where-Object { $_.Name -eq $SheetName } | Select-Object -First 1
    # Kein Rueckfall auf das erste Sheet: das importierte sonst lautlos ein
    # fremdes Blatt als Einstellungen. Der Aufrufer faengt die Exception ab.
    if (-not $sheet) { throw "Sheet '$SheetName' wurde in '$([System.IO.Path]::GetFileName($Path))' nicht gefunden." }

    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $shared = Get-XlsxSharedString -Zip $zip -MainNs $mainNs

        $shEntry = $zip.Entries | Where-Object { $_.FullName -eq $sheet.Path } | Select-Object -First 1
        if (-not $shEntry) { return $result }

        $shDoc = New-Object System.Xml.XmlDocument
        $s = $shEntry.Open(); $shDoc.Load($s); $s.Dispose()
        $rns = New-Object System.Xml.XmlNamespaceManager($shDoc.NameTable)
        $rns.AddNamespace('m', $mainNs)

        $first = $true
        foreach ($row in $shDoc.SelectNodes('//m:sheetData/m:row', $rns)) {
            $vals = @{}
            $auto = 0
            foreach ($c in $row.SelectNodes('m:c', $rns)) {
                $idx = Get-ColumnIndexFromRef $c.GetAttribute('r')
                if ($idx -lt 0) { $idx = $auto }
                $auto = $idx + 1
                $val = Get-XlsxCellValue -Cell $c -Shared $shared -Nsm $rns
                $vals[$idx] = $val
            }
            if ($first) {
                $first = $false
                $h0 = ''; $h1 = ''
                if ($vals.ContainsKey(0)) { $h0 = ([string]$vals[0]).Trim() }
                if ($vals.ContainsKey(1)) { $h1 = ([string]$vals[1]).Trim() }
                # Nur eine echte Kopfzeile ueberspringen. Beginnt das Sheet direkt
                # mit Daten, ging die erste Einstellung vorher lautlos verloren.
                if ($h0 -eq 'Name' -and $h1 -eq 'Value') { continue }
            }
            $n = ''; $v = ''
            if ($vals.ContainsKey(0)) { $n = $vals[0] }
            if ($vals.ContainsKey(1)) { $v = $vals[1] }
            if (-not [string]::IsNullOrWhiteSpace($n)) { $result[$n.Trim()] = $v }
        }
    }
    finally { $zip.Dispose() }
    return $result
}

# ===========================================================================
#  4. Excel schreiben (minimale, gueltige .xlsx - von NPOI/UiPath lesbar)
# ===========================================================================

function Write-ZipEntryText {
    param(
        [System.IO.Compression.ZipArchive]$Archive,
        [string]$EntryName,
        [string]$Content
    )
    $entry  = $Archive.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    try {
        $enc    = New-Object System.Text.UTF8Encoding($false)
        $writer = New-Object System.IO.StreamWriter($stream, $enc)
        try { $writer.Write($Content); $writer.Flush() } finally { $writer.Dispose() }
    }
    finally { $stream.Dispose() }
}

function New-ConfigXlsx {
    <#
      Erzeugt eine Config-Datei mit einem Name/Value-Sheet.
      Rows: Array von @{ Name = ''; Value = '' }
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][object[]]$Rows,
        [string]$SheetName = 'Settings'
    )

    $ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $rr = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append("<worksheet xmlns=`"$ns`" xmlns:r=`"$rr`">")
    [void]$sb.Append("<dimension ref=`"A1:B$($Rows.Count + 1)`"/>")
    [void]$sb.Append('<sheetViews><sheetView workbookViewId="0"/></sheetViews>')
    [void]$sb.Append('<sheetFormatPr defaultRowHeight="15"/>')
    [void]$sb.Append('<cols><col min="1" max="1" width="26" customWidth="1"/><col min="2" max="2" width="52" customWidth="1"/></cols>')
    [void]$sb.Append('<sheetData>')
    [void]$sb.Append('<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Value</t></is></c></row>')

    $r = 1
    foreach ($row in $Rows) {
        $r++
        $name = [string]$row.Name
        $val  = [string]$row.Value
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        [void]$sb.Append("<row r=`"$r`">")
        [void]$sb.Append("<c r=`"A$r`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$(ConvertTo-XmlText $name)</t></is></c>")

        # Immer als Text schreiben. Eine Zahlenzelle machte aus '0815' die Zahl 815
        # (fuehrende Null verloren), und [double]::TryParse ohne InvariantCulture las
        # '1.5' auf einem de-DE-Rechner als 15. UiPath holt die Werte ohnehin per
        # .ToString() aus dem konfiguration_dict.
        [void]$sb.Append("<c r=`"B$r`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$(ConvertTo-XmlText $val)</t></is></c>")
        [void]$sb.Append('</row>')
    }
    [void]$sb.Append('</sheetData></worksheet>')
    $sheetXml = $sb.ToString()

    $contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
"@

    $rootRels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"@

    $workbook = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="$ns" xmlns:r="$rr">
<sheets><sheet name="$(ConvertTo-XmlText $SheetName)" sheetId="1" r:id="rId1"/></sheets>
</workbook>
"@

    $wbRels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"@

    $styles = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="$ns">
<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
<cellStyles count="1"><cellStyle name="Standard" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
"@

    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    # Kein Remove-Item: FileMode::Create kuerzt eine bestehende Datei bereits, und
    # das Loeschen scheiterte, wenn die Datei noch in Excel offen war.
    # $fs wird ausserhalb des inneren try geoeffnet, damit ein Fehler im
    # ZipArchive-Konstruktor das Handle nicht verwaist zuruecklaesst.
    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Create)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            Write-ZipEntryText $zip '[Content_Types].xml'        $contentTypes
            Write-ZipEntryText $zip '_rels/.rels'                $rootRels
            Write-ZipEntryText $zip 'xl/workbook.xml'            $workbook
            Write-ZipEntryText $zip 'xl/_rels/workbook.xml.rels' $wbRels
            Write-ZipEntryText $zip 'xl/styles.xml'              $styles
            Write-ZipEntryText $zip 'xl/worksheets/sheet1.xml'   $sheetXml
        }
        finally { $zip.Dispose() }
    }
    finally { $fs.Dispose() }
}

# ===========================================================================
#  5. XAML-Manipulation (UiPath-Workflows)
# ===========================================================================

$script:NsAct    = 'http://schemas.microsoft.com/netfx/2009/xaml/activities'
$script:NsUi     = 'http://schemas.uipath.com/workflow/activities'
$script:NsX      = 'http://schemas.microsoft.com/winfx/2006/xaml'
$script:NsSap    = 'http://schemas.microsoft.com/netfx/2009/xaml/activities/presentation'
$script:NsSap10  = 'http://schemas.microsoft.com/netfx/2010/xaml/activities/presentation'
$script:NsScg    = 'clr-namespace:System.Collections.Generic;assembly=System.Private.CoreLib'
$script:NsSd     = 'clr-namespace:System.Data;assembly=System.Data.Common'

function Open-XamlDocument {
    param([Parameter(Mandatory)][string]$Path)
    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $true
    $doc.Load($Path)
    return $doc
}

function Get-XamlNamespaceManager {
    param([System.Xml.XmlDocument]$Doc)
    $nsm = New-Object System.Xml.XmlNamespaceManager($Doc.NameTable)
    $nsm.AddNamespace('a',       $script:NsAct)
    $nsm.AddNamespace('ui',      $script:NsUi)
    $nsm.AddNamespace('x',       $script:NsX)
    $nsm.AddNamespace('sap',     $script:NsSap)
    $nsm.AddNamespace('sap2010', $script:NsSap10)
    $nsm.AddNamespace('scg',     $script:NsScg)
    $nsm.AddNamespace('sd',      $script:NsSd)
    # Komma verhindert, dass PowerShell den (aufzaehlbaren) NamespaceManager entrollt.
    return ,$nsm
}

function Save-XamlDocument {
    <# Speichert ohne XML-Deklaration, ohne BOM, mit CRLF - wie das UiPath-Original. #>
    param([System.Xml.XmlDocument]$Doc, [string]$Path)

    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.OmitXmlDeclaration = $true
    $settings.Indent             = $false
    # XmlDocument.Load normalisiert jedes CRLF zu LF (XML 1.0 §2.11). Ohne
    # Replace/NewLineChars kaeme die Datei LF-only heraus, waehrend alle
    # Template-XAMLs CRLF sind - jede Zeile erschiene im Diff als geaendert.
    $settings.NewLineHandling    = [System.Xml.NewLineHandling]::Replace
    $settings.NewLineChars       = "`r`n"
    $settings.Encoding           = New-Object System.Text.UTF8Encoding($false)

    $sw = [System.Xml.XmlWriter]::Create($Path, $settings)
    try { $Doc.Save($sw) } finally { $sw.Dispose() }
}

# ===========================================================================
#  6. Generierung
# ===========================================================================

# Feste Infrastruktur-Variablen von PopulateQueue.xaml. Script-scope, weil
# ConvertTo-VbIdentifier diese Namen kennen muss: eine Excel-Spalte 'Reference'
# erzeugte sonst ein zweites <Variable Name="Reference" /> und Studio
# verweigerte das Oeffnen ("Variable is already declared in this scope").
$script:KeepVars = @(
    'queueName', 'primaerschluesselSpalte', 'datumsStempel', 'erstellZeitstempel',
    'zeilenZaehler', 'neueSpalte', 'chunkGroesse', 'chunk_liste', 'fehlerZeilen_dt',
    'anzErfolgreich', 'anzDuplikate', 'Reference', 'qItemInfo'
)

function Get-ActiveField {
    param($Fields)
    return @($Fields | Where-Object { $_.Selected -and -not [string]::IsNullOrWhiteSpace($_.QueueKey) })
}

function New-ReferenceExpression {
    <# Baut den VB-Ausdruck fuer die Queue-Item-Reference aus den Bausteinen. #>
    param([object[]]$Blocks, [bool]$TrimParts)

    if (-not $Blocks -or $Blocks.Count -eq 0) { return $null }
    $parts = New-Object System.Collections.ArrayList
    foreach ($b in $Blocks) {
        $e = $b.Expression
        if ($TrimParts -and $b.Kind -eq 'Feld') { $e = "$e.Trim" }
        [void]$parts.Add($e)
    }
    return ($parts -join ' & ')
}

function New-QItemDictionary {
    <# Erzeugt das VB-Dictionary-Literal fuer qItemInfo. #>
    param([object[]]$Fields, [string]$Indent = "`t")

    $lines = New-Object System.Collections.ArrayList
    foreach ($f in $Fields) {
        $key = ConvertTo-VbLiteral $f.QueueKey
        $val = $null
        if ($f.Kind -eq 'Auto') {
            if ($f.VarName -eq '@@CREATETS@@')      { $val = 'DateTime.Now' }
            elseif ($f.VarName -eq '@@FILENAME@@')  { $val = 'in_dokumentname' }
        }
        if (-not $val) { $val = $f.VarName }
        [void]$lines.Add("$Indent{$key, $val}")
    }
    return "[New Dictionary(Of String, Object) From {`r`n" + ($lines -join ",`r`n") + "`r`n}]"
}

function Update-ProjectJson {
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Path,
        [string]$ProjectName,
        [string]$Description
    )

    $json = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    $json.name        = $ProjectName
    $json.description = $Description
    $json.projectId   = [guid]::NewGuid().ToString()
    if ($json.entryPoints) {
        foreach ($ep in $json.entryPoints) { $ep.uniqueId = [guid]::NewGuid().ToString() }
    }
    $out = $json | ConvertTo-Json -Depth 20
    [System.IO.File]::WriteAllText($Path, $out, (New-Object System.Text.UTF8Encoding($false)))
}

function Update-MainXaml {
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Path,
        [string]$DefaultEnvironment,
        [string]$UtilityFolderName,
        [string]$DataFolderName,
        [string]$ConfigSheetName
    )

    $doc = Open-XamlDocument $Path
    $nsm = Get-XamlNamespaceManager $doc
    $changes = New-Object System.Collections.ArrayList

    # --- Default-Wert des Arguments In_Umgebung ---
    $root = $doc.DocumentElement
    foreach ($attr in @($root.Attributes)) {
        if ($attr.LocalName -like '*In_Umgebung*') {
            $attr.Value = $DefaultEnvironment
            [void]$changes.Add("Main.xaml: Default In_Umgebung = $DefaultEnvironment")
        }
    }

    # --- Pfade der InvokeWorkflowFile-Aktivitaeten korrigieren ---
    foreach ($inv in $doc.SelectNodes('//ui:InvokeWorkflowFile', $nsm)) {
        $wf = $inv.GetAttribute('WorkflowFileName')
        if ($wf) {
            $leaf = ($wf -split '\\')[-1]
            $new  = "$UtilityFolderName\$leaf"
            if ($new -ne $wf) {
                $inv.SetAttribute('WorkflowFileName', $new)
                [void]$changes.Add("Main.xaml: Invoke-Pfad '$wf' -> '$new'")
            }
        }

        # --- Config-Argumente von InitAllSettings ---
        foreach ($arg in $inv.SelectNodes('ui:InvokeWorkflowFile.Arguments/*', $nsm)) {
            $key = $arg.GetAttribute('Key', $script:NsX)
            if ($key -eq 'in_ConfigFile') {
                # Ordnername kommt aus einem freien Textfeld -> als VB-Literal escapen.
                $cfgPattern = ConvertTo-VbLiteral ("$DataFolderName\Config_{0}.xlsx")
                $arg.InnerText = "[String.Format($cfgPattern, umgebung)]"
                [void]$changes.Add("Main.xaml: in_ConfigFile -> $DataFolderName\Config_{DEV|PRD}.xlsx")
            }
            elseif ($key -eq 'in_ConfigSheets') {
                $arg.InnerText = $ConfigSheetName
                [void]$changes.Add("Main.xaml: in_ConfigSheets -> $ConfigSheetName")
            }
        }
    }

    Save-XamlDocument $doc $Path
    return $changes
}

function Update-ReadSourceDataXaml {
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Path, [string]$FileFilter)

    $doc = Open-XamlDocument $Path
    $nsm = Get-XamlNamespaceManager $doc
    $changes = New-Object System.Collections.ArrayList

    # --- Dateifilter setzen ---
    $forEachFile = $doc.SelectSingleNode('//ui:ForEachFileX', $nsm)
    if ($forEachFile) {
        $forEachFile.SetAttribute('Filter', $FileFilter)
        [void]$changes.Add("ReadSourceData.xaml: Dateifilter = $FileFilter")
    }

    # --- out_dokumentname wird im Template nie zugewiesen -> ergaenzen ---
    $body = $doc.SelectSingleNode('//ui:ForEachFileX/ui:ForEachFileX.Body/a:ActivityAction/a:Sequence', $nsm)
    if ($body) {
        $already = $false
        foreach ($n in $body.SelectNodes('a:Assign/a:Assign.To/a:OutArgument', $nsm)) {
            if ($n.InnerText -match 'out_dokumentname') { $already = $true }
        }
        if (-not $already) {
            $anchor = $body.SelectSingleNode('a:Assign', $nsm)
            $newEl  = $doc.CreateElement('Assign', $script:NsAct)
            if ($anchor) { [void]$body.InsertBefore($newEl, $anchor) } else { [void]$body.AppendChild($newEl) }
            $newEl.SetAttribute('DisplayName', 'Assign - Dokumentname setzen')
            $newEl.InnerXml = @'
<Assign.To><OutArgument x:TypeArguments="x:String">[out_dokumentname]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments="x:String">[CurrentFile.Name]</InArgument></Assign.Value>
'@
            [void]$changes.Add('ReadSourceData.xaml: fehlende Zuweisung out_dokumentname = CurrentFile.Name ergaenzt')
        }
    }

    Save-XamlDocument $doc $Path
    return $changes
}

function Update-PopulateQueueXaml {
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Path,
        [object[]]$Fields,
        [string]$ReferenceExpression,
        [bool]$TruncateReference,
        [string]$Priority,
        [string]$QueueNameKey,
        [string]$QueueFolderKey
    )

    $doc = Open-XamlDocument $Path
    $nsm = Get-XamlNamespaceManager $doc
    $changes = New-Object System.Collections.ArrayList

    $rootSeq = $doc.SelectSingleNode('/a:Activity/a:Sequence', $nsm)
    if (-not $rootSeq) { throw 'PopulateQueue.xaml: Haupt-Sequence nicht gefunden.' }

    # --- 1. Alte Business-Variablen ermitteln (fuer spaeteres Aufraeumen) ---
    $keep = $script:KeepVars
    $varsNode = $rootSeq.SelectSingleNode('a:Sequence.Variables', $nsm)
    if (-not $varsNode) {
        $varsNode = $doc.CreateElement('Sequence.Variables', $script:NsAct)
        [void]$rootSeq.PrependChild($varsNode)
    }
    $newNames = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($f in $Fields) { if ($f.VarName) { [void]$newNames.Add($f.VarName.ToLowerInvariant()) } }

    $oldBusinessVars = New-Object System.Collections.ArrayList
    foreach ($v in $varsNode.SelectNodes('a:Variable', $nsm)) {
        $n = $v.GetAttribute('Name')
        if (-not $n) { continue }
        if ($keep -contains $n) { continue }
        # Namen, die im neuen Dispatcher wiederverwendet werden, duerfen nicht bereinigt werden.
        if ($newNames.Contains($n.ToLowerInvariant())) { continue }
        [void]$oldBusinessVars.Add($n)
    }

    # --- 2. Variablenliste neu aufbauen ---
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("`r`n")
    foreach ($k in $keep) {
        switch ($k) {
            'zeilenZaehler'   { [void]$sb.Append("      <Variable x:TypeArguments=`"x:Int32`" Default=`"2`" Name=`"zeilenZaehler`" />`r`n") }
            'chunkGroesse'    { [void]$sb.Append("      <Variable x:TypeArguments=`"x:Int32`" Default=`"10000`" Name=`"chunkGroesse`" />`r`n") }
            'anzErfolgreich'  { [void]$sb.Append("      <Variable x:TypeArguments=`"x:Int32`" Default=`"0`" Name=`"anzErfolgreich`" />`r`n") }
            'anzDuplikate'    { [void]$sb.Append("      <Variable x:TypeArguments=`"x:Int32`" Default=`"0`" Name=`"anzDuplikate`" />`r`n") }
            'neueSpalte'      { [void]$sb.Append("      <Variable x:TypeArguments=`"sd:DataColumn`" Name=`"neueSpalte`" />`r`n") }
            'chunk_liste'     { [void]$sb.Append("      <Variable x:TypeArguments=`"scg:List(sd:DataTable)`" Name=`"chunk_liste`" />`r`n") }
            'fehlerZeilen_dt' { [void]$sb.Append("      <Variable x:TypeArguments=`"sd:DataTable`" Name=`"fehlerZeilen_dt`" />`r`n") }
            'qItemInfo'       { [void]$sb.Append("      <Variable x:TypeArguments=`"scg:Dictionary(x:String, x:Object)`" Name=`"qItemInfo`" />`r`n") }
            default           { [void]$sb.Append("      <Variable x:TypeArguments=`"x:String`" Name=`"$k`" />`r`n") }
        }
    }
    foreach ($f in $Fields) {
        if ($f.Kind -eq 'Auto') { continue }
        [void]$sb.Append("      <Variable x:TypeArguments=`"x:String`" Name=`"$($f.VarName)`" />`r`n")
    }
    [void]$sb.Append('    ')
    $varsNode.InnerXml = $sb.ToString()
    [void]$changes.Add("PopulateQueue.xaml: $((@($Fields | Where-Object { $_.Kind -ne 'Auto' })).Count) Feld-Variablen neu erzeugt")

    # --- 3. MultipleAssign neu aufbauen ---
    $multi = $doc.SelectSingleNode('//ui:MultipleAssign', $nsm)
    if (-not $multi) { throw 'PopulateQueue.xaml: MultipleAssign-Aktivitaet nicht gefunden.' }
    $opsHost = $multi.SelectSingleNode('ui:MultipleAssign.AssignOperations', $nsm)
    if (-not $opsHost) { throw 'PopulateQueue.xaml: MultipleAssign.AssignOperations nicht gefunden.' }

    $assignFields = @($Fields | Where-Object { $_.Kind -ne 'Auto' })
    $ob = New-Object System.Text.StringBuilder
    [void]$ob.Append("`r`n          <scg:List x:TypeArguments=`"ui:AssignOperation`" Capacity=`"$([Math]::Max($assignFields.Count,1))`">`r`n")
    $i = 0
    foreach ($f in $assignFields) {
        $i++
        $value = switch ($f.Kind) {
            'Spalte'   { '[CurrentRow(' + (ConvertTo-VbLiteral $f.ColumnName) + ').ToString]' }
            'Text'     { '[' + (ConvertTo-VbLiteral $f.Expression) + ']' }
            'Ausdruck' { '[(' + $f.Expression + ').ToString]' }
            default    { '[String.Empty]' }
        }
        [void]$ob.Append("            <ui:AssignOperation sap2010:WorkflowViewState.IdRef=`"AssignOperation_$i`">`r`n")
        [void]$ob.Append("              <ui:AssignOperation.To>`r`n")
        [void]$ob.Append("                <OutArgument x:TypeArguments=`"x:String`">[$($f.VarName)]</OutArgument>`r`n")
        [void]$ob.Append("              </ui:AssignOperation.To>`r`n")
        [void]$ob.Append("              <ui:AssignOperation.Value>`r`n")
        [void]$ob.Append("                <InArgument x:TypeArguments=`"x:String`">$(ConvertTo-XmlText $value)</InArgument>`r`n")
        [void]$ob.Append("              </ui:AssignOperation.Value>`r`n")
        [void]$ob.Append("            </ui:AssignOperation>`r`n")
    }
    [void]$ob.Append("          </scg:List>`r`n        ")
    $opsHost.InnerXml = $ob.ToString()
    $multi.SetAttribute('DisplayName', 'Multiple Assign - Daten pro Zeile aus DataTable auslesen')
    [void]$changes.Add("PopulateQueue.xaml: MultipleAssign mit $($assignFields.Count) Zuweisungen neu erzeugt")

    # --- 4. qItemInfo-Dictionary ---
    $qAssign = $null
    foreach ($a in $doc.SelectNodes('//a:Assign', $nsm)) {
        if ($a.GetAttribute('DisplayName') -eq 'Assign QItemInfo') { $qAssign = $a; break }
    }
    if (-not $qAssign) {
        foreach ($a in $doc.SelectNodes('//a:Assign', $nsm)) {
            $to = $a.SelectSingleNode('a:Assign.To/a:OutArgument', $nsm)
            if ($to -and $to.InnerText -match 'qItemInfo') { $qAssign = $a; break }
        }
    }
    if (-not $qAssign) { throw 'PopulateQueue.xaml: Assign fuer qItemInfo nicht gefunden.' }

    $inArg = $qAssign.SelectSingleNode('a:Assign.Value/a:InArgument', $nsm)
    if ($inArg) {
        $inArg.InnerText = New-QItemDictionary -Fields $Fields
        $sp = $inArg.OwnerDocument.CreateAttribute('xml', 'space', 'http://www.w3.org/XML/1998/namespace')
        $sp.Value = 'preserve'
        [void]$inArg.Attributes.SetNamedItem($sp)
        [void]$changes.Add("PopulateQueue.xaml: qItemInfo-Dictionary mit $($Fields.Count) Eintraegen erzeugt")
    }

    # --- 5. Reference-Assign vor dem qItemInfo-Assign einfuegen ---
    $bodySeq = $qAssign.ParentNode
    foreach ($old in @($bodySeq.SelectNodes('a:Assign', $nsm))) {
        if ($old.GetAttribute('DisplayName') -like 'Assign - Reference*') { [void]$bodySeq.RemoveChild($old) }
    }

    $refAssign = $doc.CreateElement('Assign', $script:NsAct)
    [void]$bodySeq.InsertBefore($refAssign, $qAssign)
    $refAssign.SetAttribute('DisplayName', 'Assign - Reference erstellen')
    $refAssign.InnerXml = "<Assign.To><OutArgument x:TypeArguments=`"x:String`">[Reference]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=`"x:String`">$(ConvertTo-XmlText "[$ReferenceExpression]")</InArgument></Assign.Value>"
    [void]$changes.Add("PopulateQueue.xaml: Reference = $ReferenceExpression")

    if ($TruncateReference) {
        $cut = $doc.CreateElement('Assign', $script:NsAct)
        [void]$bodySeq.InsertBefore($cut, $qAssign)
        $cut.SetAttribute('DisplayName', 'Assign - Reference auf 128 Zeichen kuerzen')
        $cut.InnerXml = "<Assign.To><OutArgument x:TypeArguments=`"x:String`">[Reference]</OutArgument></Assign.To><Assign.Value><InArgument x:TypeArguments=`"x:String`">$(ConvertTo-XmlText '[If(Reference.Length > 128, Reference.Substring(0, 128), Reference)]')</InArgument></Assign.Value>"
        [void]$changes.Add('PopulateQueue.xaml: Reference wird auf 128 Zeichen begrenzt (Orchestrator-Limit)')
    }

    # --- 6. AddQueueItem konfigurieren ---
    # Beide Schluessel stammen aus freien Textfeldern der Maske 4. Ein Anfuehrungs-
    # zeichen darin ergab bisher einen VB-Syntaxfehler, den erst Studio meldete.
    $qNameLit   = ConvertTo-VbLiteral $QueueNameKey
    $qFolderLit = ConvertTo-VbLiteral $QueueFolderKey

    $add = $doc.SelectSingleNode('//ui:AddQueueItem', $nsm)
    if ($add) {
        $add.SetAttribute('Reference', '[Reference]')
        $add.SetAttribute('Priority', $Priority)
        $add.SetAttribute('QueueType',  "[in_konfiguration_dict($qNameLit).ToString]")
        $add.SetAttribute('FolderPath', "[in_konfiguration_dict($qFolderLit).ToString]")
        $add.SetAttribute('ItemInformationCollection', '[qItemInfo]')
        [void]$changes.Add("PopulateQueue.xaml: AddQueueItem -> Reference=[Reference], Prioritaet=$Priority, Queue=$QueueNameKey, Folder=$QueueFolderKey")
    }

    # --- 7. Toten Comment-Out-Block entfernen (verweist auf nicht existierende Variablen) ---
    foreach ($co in @($doc.SelectNodes('//ui:CommentOut', $nsm))) {
        [void]$co.ParentNode.RemoveChild($co)
        [void]$changes.Add('PopulateQueue.xaml: veralteter Comment-Out-Block entfernt')
    }

    # --- 8a. Start-Logmeldung auf den konfigurierten Queue-Schluessel umstellen ---
    foreach ($log in $doc.SelectNodes('//ui:LogMessage', $nsm)) {
        if ($log.GetAttribute('DisplayName') -eq 'Log: Start') {
            $log.SetAttribute('Message', "[""Start | PopulateQueue | Queue="" & in_konfiguration_dict($qNameLit).ToString]")
        }
    }

    # --- 8b. Logmeldungen auf die neue Reference umstellen ---
    foreach ($log in $doc.SelectNodes('//ui:LogMessage', $nsm)) {
        $msg = $log.GetAttribute('Message')
        if (-not $msg) { continue }
        $hit = $false
        foreach ($ov in $oldBusinessVars) { if ($msg -match [regex]::Escape($ov)) { $hit = $true; break } }
        if ($hit) {
            if ($log.GetAttribute('Level') -eq 'Error') {
                $log.SetAttribute('Message', '["Fehler beim Hochladen der Zeile | Reference: " & Reference & " | Fehler: " & exception.Message]')
            }
            else {
                $log.SetAttribute('Message', '["Abgeschlossen | PopulateQueue | Reference: " & Reference]')
            }
            [void]$changes.Add('PopulateQueue.xaml: Logmeldung auf Reference umgestellt')
        }
    }

    # --- 9. Auf verbliebene Referenzen zu alten Fach-Variablen hinweisen ---
    # Hier wurde frueher jeder Treffer im gesamten OuterXml per Regex durch
    # 'Reference' ersetzt. Das traf auch die soeben erzeugten VB-Literale:
    # QueueKey und ColumnName stehen ungefaltet im XML, ihr VarName dagegen
    # umlautgefaltet (ae statt ae). Eine Spalte 'KontounabhaengigeZeile...' galt
    # damit als alter Bezeichner, und [CurrentRow("Kontounabh...")] wurde zu
    # [CurrentRow("Reference")] - zur Laufzeit "Column does not belong to table",
    # bei zwei solchen Spalten zusaetzlich ein doppelter Dictionary-Schluessel.
    # Ersetzt wird deshalb nichts mehr; gemeldet wird nur noch.
    if ($oldBusinessVars.Count -gt 0) {
        $expr = New-Object System.Text.StringBuilder
        foreach ($n in $doc.SelectNodes('//a:InArgument | //a:OutArgument', $nsm)) { [void]$expr.AppendLine($n.InnerText) }
        foreach ($n in $doc.SelectNodes('//ui:LogMessage/@Message', $nsm))          { [void]$expr.AppendLine($n.Value) }
        $haystack = $expr.ToString()
        foreach ($ov in $oldBusinessVars) {
            if ($haystack -match "\b$([regex]::Escape($ov))\b") {
                [void]$changes.Add("PopulateQueue.xaml: ACHTUNG - alter Bezeichner '$ov' wird noch in einem Ausdruck verwendet, ist aber nicht mehr deklariert. Bitte in Studio pruefen.")
            }
        }
    }

    Save-XamlDocument $doc $Path
    return $changes
}

function Invoke-DispatcherGeneration {
    param([hashtable]$Cfg)

    $log = New-Object System.Collections.ArrayList
    $tpl = $Cfg.TemplatePath
    $out = $Cfg.OutputPath

    if (-not (Test-Path -LiteralPath $tpl)) { throw "Template-Ordner nicht gefunden: $tpl" }

    # --- Ziel gegen Template pruefen ---
    # Ohne diese Pruefung kopierte der Generator das Template in sich selbst und
    # schrieb danach project.json, Main.xaml, PopulateQueue.xaml und beide
    # Config-Dateien des Templates um. Es gibt kein Backup, das war endgueltig.
    $sep     = [System.IO.Path]::DirectorySeparatorChar
    $tplFull = [System.IO.Path]::GetFullPath($tpl).TrimEnd($sep)
    $outFull = [System.IO.Path]::GetFullPath($out).TrimEnd($sep)
    if ($outFull.Equals($tplFull, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Ziel-Ordner und Template-Ordner sind identisch. Bitte einen anderen Ziel-Ordner waehlen."
    }
    if ($outFull.StartsWith($tplFull + $sep, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Der Ziel-Ordner liegt innerhalb des Template-Ordners. Bitte einen Ordner ausserhalb des Templates waehlen."
    }

    # --- Zielordner vorbereiten ---
    if (-not (Test-Path -LiteralPath $out)) {
        New-Item -ItemType Directory -Path $out -Force | Out-Null
        [void]$log.Add("Zielordner erstellt: $out")
    }

    # --- Template kopieren ---
    $skip = @('.local', '.settings', '.objects', '.entities', '.tmh', '.git', '.vs', '.screenshots')
    $selfName = 'New-UiPathDispatcher.ps1'   # der Generator selbst wandert nie ins neue Projekt
    $items = @(Get-ChildItem -LiteralPath $tpl -Force |
        Where-Object { $skip -notcontains $_.Name -and $_.Name -ne '.DS_Store' -and $_.Name -ne $selfName })
    foreach ($item in $items) {
        $dest = Join-Path $out $item.Name
        if ($item.PSIsContainer) {
            # Copy-Item -Recurse haengt den Quellordnernamen erneut an, wenn das Ziel
            # bereits existiert -> Out\Utility_Files\Utility_Files. Beim zweiten Lauf
            # in denselben Ordner bearbeitete der Generator danach die alte, von
            # UiPath nie ausgefuehrte PopulateQueue.xaml und meldete trotzdem Erfolg.
            if (Test-Path -LiteralPath $dest) { Remove-Item -LiteralPath $dest -Recurse -Force -ErrorAction Stop }
            Copy-Item -LiteralPath $item.FullName -Destination $dest -Recurse -Force -ErrorAction Stop
        }
        else {
            # -ErrorAction Stop: ein gesperrte Datei wurde vorher stillschweigend
            # uebersprungen und der Lauf endete mit "ohne Fehler abgeschlossen".
            Copy-Item -LiteralPath $item.FullName -Destination $dest -Force -ErrorAction Stop
        }
    }
    Get-ChildItem -LiteralPath $out -Recurse -Force -Filter '.DS_Store' -ErrorAction SilentlyContinue |
        Remove-Item -Force -ErrorAction SilentlyContinue
    [void]$log.Add("Template kopiert: $tpl -> $out")

    # --- Utility-Ordner + Dateien lokalisieren ---
    $popFile = Get-ChildItem -LiteralPath $out -Recurse -Filter 'PopulateQueue.xaml' | Select-Object -First 1
    if (-not $popFile) { throw 'PopulateQueue.xaml wurde im kopierten Projekt nicht gefunden.' }
    $utilDir  = $popFile.Directory
    $utilName = $utilDir.Name

    $mainFile = Join-Path $out 'Main.xaml'
    $readFile = Join-Path $utilDir.FullName 'ReadSourceData.xaml'
    $projFile = Join-Path $out 'project.json'

    # --- project.json ---
    if (Test-Path -LiteralPath $projFile) {
        Update-ProjectJson -Path $projFile -ProjectName $Cfg.ProjectName -Description $Cfg.ProjectDescription
        [void]$log.Add("project.json: Name = '$($Cfg.ProjectName)', neue Projekt-GUIDs vergeben")
    }

    # --- Main.xaml ---
    if (Test-Path -LiteralPath $mainFile) {
        $c = Update-MainXaml -Path $mainFile `
                             -DefaultEnvironment $Cfg.DefaultEnvironment `
                             -UtilityFolderName $utilName `
                             -DataFolderName $Cfg.DataFolderName `
                             -ConfigSheetName $Cfg.ConfigSheetName
        foreach ($x in $c) { [void]$log.Add($x) }
    }

    # --- ReadSourceData.xaml ---
    if (Test-Path -LiteralPath $readFile) {
        $c = Update-ReadSourceDataXaml -Path $readFile -FileFilter $Cfg.FileFilter
        foreach ($x in $c) { [void]$log.Add($x) }
    }

    # --- PopulateQueue.xaml ---
    $c = Update-PopulateQueueXaml -Path $popFile.FullName `
                                  -Fields $Cfg.Fields `
                                  -ReferenceExpression $Cfg.ReferenceExpression `
                                  -TruncateReference $Cfg.TruncateReference `
                                  -Priority $Cfg.Priority `
                                  -QueueNameKey $Cfg.QueueNameKey `
                                  -QueueFolderKey $Cfg.QueueFolderKey
    foreach ($x in $c) { [void]$log.Add($x) }

    # --- Config-Dateien ---
    $dataDir = Join-Path $out $Cfg.DataFolderName
    if (-not (Test-Path -LiteralPath $dataDir)) { New-Item -ItemType Directory -Path $dataDir -Force | Out-Null }

    $devRows = @(); $prdRows = @()
    foreach ($ci in $Cfg.ConfigItems) {
        if ([string]::IsNullOrWhiteSpace($ci.Name)) { continue }
        $devRows += ,@{ Name = $ci.Name; Value = $ci.DevValue }
        $prdRows += ,@{ Name = $ci.Name; Value = $ci.PrdValue }
    }

    $devPath = Join-Path $dataDir 'Config_DEV.xlsx'
    $prdPath = Join-Path $dataDir 'Config_PRD.xlsx'
    New-ConfigXlsx -Path $devPath -Rows $devRows -SheetName $Cfg.ConfigSheetName
    New-ConfigXlsx -Path $prdPath -Rows $prdRows -SheetName $Cfg.ConfigSheetName
    [void]$log.Add("Config_DEV.xlsx geschrieben ($($devRows.Count) Eintraege)")
    [void]$log.Add("Config_PRD.xlsx geschrieben ($($prdRows.Count) Eintraege)")

    return $log
}

# ===========================================================================
#  7. GUI (WPF / XAML)
# ===========================================================================

$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Dispatcher Generator" Height="900" Width="1280"
        MinWidth="1040" MinHeight="720"
        WindowStartupLocation="CenterScreen" Background="#FFF3F1F0">

  <!-- ===================================================================
       Corporate Design Schwyzer Kantonalbank
       Brand Design Manual V3.1 / Juni 2026
         Primaerfarben    Weiss #FFFFFF, SZKB Rot #E6181A,
                          SZKB Mittelrot #AE0B04, SZKB Dunkelrot #770A00,
                          Schwarz #000000  (duerfen nicht aufgehellt werden)
         Sekundaerfarben  SZKB Grau #92847B / 50% #C5BDB6 / 35% #D6CFCB /
                          20% #E7E3E0 / 10% #F3F1F0
         Tertiaerfarben   Hellblau #96ACB6, Blau #5C6A99, Hellviolett #B5A1B4,
                          Violett #492745, Leuchtrot #E84025
         Bilanzfarben     Gruen #5A9D28, Orange #E89C00 (nur Status/Bilanz)
         Typografie       Futura Now (Hauptschrift), intern Futura Now Text
         Gestaltungselement «Wegweiser» - immer angeschnitten, max. 2 Elemente,
                          nie gedreht oder gespiegelt, nicht zu maechtig
       =================================================================== -->

  <Window.Resources>

    <FontFamily x:Key="BrandFont">Futura Now Text, Futura Now, Futura, Century Gothic, Segoe UI</FontFamily>
    <FontFamily x:Key="MonoFont">Consolas, Courier New</FontFamily>

    <SolidColorBrush x:Key="Weiss"         Color="#FFFFFFFF"/>
    <SolidColorBrush x:Key="Schwarz"       Color="#FF000000"/>
    <SolidColorBrush x:Key="SzkbRot"       Color="#FFE6181A"/>
    <SolidColorBrush x:Key="SzkbMittelrot" Color="#FFAE0B04"/>
    <SolidColorBrush x:Key="SzkbDunkelrot" Color="#FF770A00"/>
    <SolidColorBrush x:Key="SzkbGrau"      Color="#FF92847B"/>
    <SolidColorBrush x:Key="Grau50"        Color="#FFC5BDB6"/>
    <SolidColorBrush x:Key="Grau35"        Color="#FFD6CFCB"/>
    <SolidColorBrush x:Key="Grau20"        Color="#FFE7E3E0"/>
    <SolidColorBrush x:Key="Grau10"        Color="#FFF3F1F0"/>
    <SolidColorBrush x:Key="Hellblau"      Color="#FF96ACB6"/>
    <SolidColorBrush x:Key="BilanzGruen"   Color="#FF5A9D28"/>

    <!-- ========== Text ========== -->
    <Style TargetType="TextBlock">
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Foreground" Value="{StaticResource Schwarz}"/>
      <Setter Property="FontFamily" Value="{StaticResource BrandFont}"/>
      <Setter Property="FontSize" Value="13"/>
    </Style>

    <Style x:Key="Section" TargetType="TextBlock">
      <Setter Property="FontFamily" Value="{StaticResource BrandFont}"/>
      <Setter Property="FontWeight" Value="Bold"/>
      <Setter Property="FontSize" Value="15"/>
      <Setter Property="Margin" Value="0,12,0,4"/>
      <Setter Property="Foreground" Value="{StaticResource Schwarz}"/>
    </Style>

    <Style x:Key="Hint" TargetType="TextBlock">
      <Setter Property="FontFamily" Value="{StaticResource BrandFont}"/>
      <Setter Property="Foreground" Value="{StaticResource SzkbGrau}"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="Margin" Value="0,2,0,10"/>
    </Style>

    <Style x:Key="Caption" TargetType="TextBlock">
      <Setter Property="FontFamily" Value="{StaticResource BrandFont}"/>
      <Setter Property="Foreground" Value="{StaticResource SzkbGrau}"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="FontWeight" Value="Bold"/>
      <Setter Property="Margin" Value="0,0,0,4"/>
    </Style>

    <!-- ========== Eingabefelder ========== -->
    <Style TargetType="TextBox">
      <Setter Property="Height" Value="28"/>
      <Setter Property="Margin" Value="0,4,0,4"/>
      <Setter Property="Padding" Value="7,0,7,0"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="FontFamily" Value="{StaticResource BrandFont}"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Background" Value="{StaticResource Weiss}"/>
      <Setter Property="Foreground" Value="{StaticResource Schwarz}"/>
      <Setter Property="BorderBrush" Value="{StaticResource Grau50}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="SelectionBrush" Value="{StaticResource SzkbRot}"/>
      <Style.Triggers>
        <Trigger Property="IsKeyboardFocusWithin" Value="True">
          <Setter Property="BorderBrush" Value="{StaticResource SzkbRot}"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <Style TargetType="ComboBox">
      <Setter Property="Height" Value="28"/>
      <Setter Property="Margin" Value="0,4,0,4"/>
      <Setter Property="FontFamily" Value="{StaticResource BrandFont}"/>
      <Setter Property="FontSize" Value="13"/>
    </Style>

    <Style TargetType="CheckBox">
      <Setter Property="FontFamily" Value="{StaticResource BrandFont}"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Foreground" Value="{StaticResource Schwarz}"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>

    <!-- ========== Buttons ========== -->
    <Style TargetType="Button">
      <Setter Property="Height" Value="32"/>
      <Setter Property="MinWidth" Value="96"/>
      <Setter Property="Margin" Value="6,0,0,0"/>
      <Setter Property="Padding" Value="16,0,16,0"/>
      <Setter Property="FontFamily" Value="{StaticResource BrandFont}"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Foreground" Value="{StaticResource Schwarz}"/>
      <Setter Property="Background" Value="{StaticResource Weiss}"/>
      <Setter Property="BorderBrush" Value="{StaticResource Grau50}"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" SnapsToDevicePixels="True"
                    Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"
                                Margin="{TemplateBinding Padding}"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{StaticResource Grau20}"/>
                <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource SzkbGrau}"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{StaticResource Grau35}"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="bd" Property="Opacity" Value="0.4"/>
                <Setter Property="Cursor" Value="Arrow"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="BtnPrimary" TargetType="Button">
      <Setter Property="Height" Value="32"/>
      <Setter Property="MinWidth" Value="110"/>
      <Setter Property="Margin" Value="6,0,0,0"/>
      <Setter Property="Padding" Value="18,0,18,0"/>
      <Setter Property="FontFamily" Value="{StaticResource BrandFont}"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="FontWeight" Value="Bold"/>
      <Setter Property="Foreground" Value="{StaticResource Weiss}"/>
      <Setter Property="Background" Value="{StaticResource SzkbRot}"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" SnapsToDevicePixels="True"
                    Background="{TemplateBinding Background}" BorderThickness="0">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"
                                Margin="{TemplateBinding Padding}"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{StaticResource SzkbMittelrot}"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{StaticResource SzkbDunkelrot}"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="bd" Property="Background" Value="{StaticResource Grau50}"/>
                <Setter Property="Cursor" Value="Arrow"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ========== Tabs ========== -->
    <Style TargetType="TabItem">
      <Setter Property="FontFamily" Value="{StaticResource BrandFont}"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Foreground" Value="{StaticResource SzkbGrau}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TabItem">
            <Border x:Name="bd" Background="Transparent" Padding="20,11,20,11"
                    BorderThickness="0,0,0,3" BorderBrush="Transparent" SnapsToDevicePixels="True">
              <ContentPresenter ContentSource="Header" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{StaticResource Grau20}"/>
              </Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{StaticResource Weiss}"/>
                <Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource SzkbRot}"/>
                <Setter Property="Foreground" Value="{StaticResource Schwarz}"/>
                <Setter Property="FontWeight" Value="Bold"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="TabControl">
      <Setter Property="Background" Value="{StaticResource Weiss}"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TabControl">
            <DockPanel>
              <Border DockPanel.Dock="Top" Background="{StaticResource Grau10}"
                      BorderBrush="{StaticResource Grau35}" BorderThickness="0,0,0,1">
                <TabPanel IsItemsHost="True" Margin="6,0,0,0"/>
              </Border>
              <Border Background="{StaticResource Weiss}">
                <ContentPresenter x:Name="PART_SelectedContentHost" ContentSource="SelectedContent"/>
              </Border>
            </DockPanel>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- ========== Listen ========== -->
    <Style TargetType="GridViewColumnHeader">
      <Setter Property="FontFamily" Value="{StaticResource BrandFont}"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="FontWeight" Value="Bold"/>
      <Setter Property="Foreground" Value="{StaticResource Schwarz}"/>
      <Setter Property="HorizontalContentAlignment" Value="Left"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="GridViewColumnHeader">
            <Grid>
              <Border Background="{StaticResource Grau20}" BorderBrush="{StaticResource Grau35}"
                      BorderThickness="0,0,1,1" Height="32" SnapsToDevicePixels="True">
                <ContentPresenter Margin="8,0,6,0" VerticalAlignment="Center"
                                  HorizontalAlignment="Left"/>
              </Border>
              <Thumb x:Name="PART_HeaderGripper" HorizontalAlignment="Right" Width="8" Cursor="SizeWE">
                <Thumb.Template>
                  <ControlTemplate TargetType="Thumb">
                    <Border Background="Transparent" Width="8"/>
                  </ControlTemplate>
                </Thumb.Template>
              </Thumb>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="ListView">
      <Setter Property="BorderBrush" Value="{StaticResource Grau35}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Background" Value="{StaticResource Weiss}"/>
      <Setter Property="FontFamily" Value="{StaticResource BrandFont}"/>
      <Setter Property="FontSize" Value="13"/>
    </Style>

    <Style TargetType="ListBox">
      <Setter Property="BorderBrush" Value="{StaticResource Grau35}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Background" Value="{StaticResource Weiss}"/>
      <Setter Property="FontFamily" Value="{StaticResource BrandFont}"/>
      <Setter Property="FontSize" Value="13"/>
    </Style>

    <!-- Zelleneingabe in Listen -->
    <Style x:Key="CellBox" TargetType="TextBox">
      <Setter Property="Height" Value="24"/>
      <Setter Property="Margin" Value="0"/>
      <Setter Property="Padding" Value="4,0,4,0"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="FontFamily" Value="{StaticResource BrandFont}"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Foreground" Value="{StaticResource Schwarz}"/>
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="SelectionBrush" Value="{StaticResource SzkbRot}"/>
    </Style>

    <!-- Editierbare Zelle: unterstrichen in SZKB Rot -->
    <Style x:Key="CellBoxEdit" TargetType="TextBox" BasedOn="{StaticResource CellBox}">
      <Setter Property="Background" Value="{StaticResource Grau10}"/>
      <Setter Property="BorderThickness" Value="0,0,0,1"/>
      <Setter Property="BorderBrush" Value="{StaticResource SzkbRot}"/>
    </Style>

    <Style x:Key="CellBoxDev" TargetType="TextBox" BasedOn="{StaticResource CellBox}">
      <Setter Property="Background" Value="{StaticResource Grau10}"/>
      <Setter Property="BorderThickness" Value="0,0,0,1"/>
      <Setter Property="BorderBrush" Value="{StaticResource Grau50}"/>
    </Style>

    <Style x:Key="CellBoxPrd" TargetType="TextBox" BasedOn="{StaticResource CellBox}">
      <Setter Property="Background" Value="{StaticResource Grau20}"/>
      <Setter Property="BorderThickness" Value="0,0,0,1"/>
      <Setter Property="BorderBrush" Value="{StaticResource SzkbRot}"/>
    </Style>

    <!-- Monospace-Ausgaben (Zusammenfassung / Protokoll) -->
    <Style x:Key="MonoBox" TargetType="TextBox">
      <Setter Property="Margin" Value="0"/>
      <Setter Property="Padding" Value="12,10,12,10"/>
      <Setter Property="VerticalContentAlignment" Value="Top"/>
      <Setter Property="IsReadOnly" Value="True"/>
      <Setter Property="AcceptsReturn" Value="True"/>
      <Setter Property="TextWrapping" Value="NoWrap"/>
      <Setter Property="FontFamily" Value="{StaticResource MonoFont}"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="Foreground" Value="{StaticResource Schwarz}"/>
      <Setter Property="Background" Value="{StaticResource Grau10}"/>
      <Setter Property="BorderBrush" Value="{StaticResource Grau35}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="SelectionBrush" Value="{StaticResource SzkbRot}"/>
      <Setter Property="VerticalScrollBarVisibility" Value="Auto"/>
      <Setter Property="HorizontalScrollBarVisibility" Value="Auto"/>
    </Style>

  </Window.Resources>

  <DockPanel>

    <!-- ================= Kopf mit Gestaltungselement «Wegweiser» ================= -->
    <Border DockPanel.Dock="Top" Background="{StaticResource Weiss}" Height="104"
            BorderBrush="{StaticResource SzkbRot}" BorderThickness="0,0,0,3">
      <Grid ClipToBounds="True">
        <!-- Wegweiser: immer angeschnitten, maximal zwei Elemente, nicht gedreht -->
        <Canvas HorizontalAlignment="Right" Width="330" ClipToBounds="True">
          <Path Fill="{StaticResource Grau20}"
                Data="M 18,-46 L 196,-24 L 232,104 L 150,196 L 26,158 L 62,54 Z"/>
          <Path Fill="{StaticResource SzkbRot}"
                Data="M 150,-52 L 336,-34 L 372,118 L 268,206 L 138,160 L 180,52 Z"/>
        </Canvas>

        <StackPanel VerticalAlignment="Center" Margin="26,0,0,0">
          <TextBlock Text="Dispatcher Generator" FontSize="26" FontWeight="Bold"
                     Foreground="{StaticResource Schwarz}"/>
          <TextBlock x:Name="txtHeaderSub" FontSize="13" Margin="0,2,0,0"
                     Foreground="{StaticResource SzkbGrau}"
                     Text="UiPath-Template kopieren und automatisch auf ein neues Quell-Excel anpassen"/>
        </StackPanel>
      </Grid>
    </Border>

    <!-- ================= Fusszeile ================= -->
    <Border DockPanel.Dock="Bottom" Background="{StaticResource Grau10}"
            BorderBrush="{StaticResource Grau35}" BorderThickness="0,1,0,0" Padding="18,12,18,12">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Grid.Column="0" VerticalAlignment="Center">
          <TextBlock x:Name="txtStatus" Text="Bereit." Foreground="{StaticResource Schwarz}"/>
          <TextBlock Text="Schwyzer Kantonalbank" FontSize="11"
                     Foreground="{StaticResource SzkbGrau}" Margin="0,1,0,0"/>
        </StackPanel>
        <StackPanel Grid.Column="1" Orientation="Horizontal">
          <Button x:Name="btnBack"   Content="Zurueck"/>
          <Button x:Name="btnNext"   Content="Weiter" Style="{StaticResource BtnPrimary}"/>
          <Button x:Name="btnCreate" Content="Dispatcher erstellen" Style="{StaticResource BtnPrimary}" MinWidth="170"
                    IsEnabled="{Binding IsChecked, ElementName=chkConfirm}"/>
          <Button x:Name="btnClose"  Content="Schliessen"/>
        </StackPanel>
      </Grid>
    </Border>

    <TabControl x:Name="tabMain">

      <!-- ================= 1. Vorlage und Quelle ================= -->
      <TabItem Header="1 · Vorlage &amp; Quelle">
        <ScrollViewer VerticalScrollBarVisibility="Auto">
        <Grid Margin="26,18,26,18">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="200"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <TextBlock Grid.Row="0" Grid.ColumnSpan="3" Style="{StaticResource Section}" Text="Dispatcher-Template"/>
          <TextBlock Grid.Row="1" Grid.ColumnSpan="3" Style="{StaticResource Hint}"
                     Text="Ordner mit Main.xaml, project.json, Data\ und dem Utility-Ordner (PopulateQueue.xaml, ReadSourceData.xaml, InitAllSettings.xaml). Das Template selbst wird nicht veraendert."/>

          <TextBlock Grid.Row="2" Grid.Column="0" Text="Template-Ordner"/>
          <TextBox   Grid.Row="2" Grid.Column="1" x:Name="txtTemplate"/>
          <Button    Grid.Row="2" Grid.Column="2" x:Name="btnTemplate" Content="Durchsuchen"/>

          <TextBlock Grid.Row="3" Grid.Column="0" Text="Ziel-Ordner (neu)"/>
          <TextBox   Grid.Row="3" Grid.Column="1" x:Name="txtOutput"/>
          <Button    Grid.Row="3" Grid.Column="2" x:Name="btnOutput" Content="Durchsuchen"/>

          <TextBlock Grid.Row="4" Grid.Column="0" Text="Projektname"/>
          <TextBox   Grid.Row="4" Grid.Column="1" x:Name="txtProjectName"/>

          <TextBlock Grid.Row="5" Grid.Column="0" Text="Beschreibung"/>
          <TextBox   Grid.Row="5" Grid.Column="1" x:Name="txtProjectDesc"/>

          <TextBlock Grid.Row="6" Grid.ColumnSpan="3" Style="{StaticResource Section}" Text="Quell-Excel (Input des Dispatchers)"/>
          <TextBlock Grid.Row="7" Grid.ColumnSpan="3" Style="{StaticResource Hint}"
                     Text="Aus dieser Datei werden Sheet-Namen und Spaltenueberschriften gelesen. Sie wird nur analysiert, nicht veraendert."/>

          <TextBlock Grid.Row="8" Grid.Column="0" Text="Excel-Datei"/>
          <TextBox   Grid.Row="8" Grid.Column="1" x:Name="txtSourceExcel"/>
          <Button    Grid.Row="8" Grid.Column="2" x:Name="btnSourceExcel" Content="Durchsuchen"/>

          <TextBlock Grid.Row="9" Grid.Column="0" Text="Sheet (Tabellenblatt)"/>
          <ComboBox  Grid.Row="9" Grid.Column="1" x:Name="cboSheet"/>
          <Button    Grid.Row="9" Grid.Column="2" x:Name="btnReadColumns" Content="Spalten einlesen"
                     Style="{StaticResource BtnPrimary}" MinWidth="150"/>

          <TextBlock Grid.Row="10" Grid.ColumnSpan="3" Style="{StaticResource Section}" Text="Projekt-Struktur"/>

          <TextBlock Grid.Row="11" Grid.Column="0" Text="Config-Unterordner"/>
          <TextBox   Grid.Row="11" Grid.Column="1" x:Name="txtDataFolder" Text="Data"/>

          <TextBlock Grid.Row="12" Grid.Column="0" Text="Config-Sheet"/>
          <TextBox   Grid.Row="12" Grid.Column="1" x:Name="txtConfigSheet" Text="Settings"/>

          <TextBlock Grid.Row="13" Grid.Column="0" Text="Default In_Umgebung"/>
          <ComboBox  Grid.Row="13" Grid.Column="1" x:Name="cboEnv">
            <ComboBoxItem Content="SZT2" IsSelected="True"/>
            <ComboBoxItem Content="SZT1"/>
            <ComboBoxItem Content="SZP"/>
          </ComboBox>
        </Grid>
        </ScrollViewer>
      </TabItem>

      <!-- ================= 2. Felder ================= -->
      <TabItem Header="2 · Felder &amp; Reihenfolge">
        <Grid Margin="26,18,26,18">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <StackPanel Grid.Row="0">
            <TextBlock Style="{StaticResource Section}" Text="Queue-Item Inhalt (Specific Data)"/>
            <TextBlock Style="{StaticResource Hint}"
                       Text="Jede Zeile wird ein Eintrag im Queue-Item-Dictionary. Die Reihenfolge in dieser Liste entspricht der Speicher-Reihenfolge. 'filename' und 'CreateTS' stehen fix zuoberst (Schluesselname aenderbar, Eintrag nicht entfernbar). Rot unterstrichene Felder sind direkt editierbar."/>
          </StackPanel>

          <ListView Grid.Row="1" x:Name="lvFields" Margin="0,4,0,10" SelectionMode="Extended">
            <ListView.View>
              <GridView>
                <GridViewColumn Header="Aktiv" Width="52">
                  <GridViewColumn.CellTemplate>
                    <DataTemplate>
                      <CheckBox IsChecked="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                                IsEnabled="{Binding IsUnlocked}" HorizontalAlignment="Center"/>
                    </DataTemplate>
                  </GridViewColumn.CellTemplate>
                </GridViewColumn>
                <GridViewColumn Header="Typ" Width="86" DisplayMemberBinding="{Binding Kind}"/>
                <GridViewColumn Header="Excel-Spalte / Feld" Width="310">
                  <GridViewColumn.CellTemplate>
                    <DataTemplate>
                      <TextBox Style="{StaticResource CellBox}"
                               Text="{Binding ColumnName, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                               IsReadOnly="{Binding IsLocked}"/>
                    </DataTemplate>
                  </GridViewColumn.CellTemplate>
                </GridViewColumn>
                <GridViewColumn Header="Queue-Referenz (Schluessel)" Width="280">
                  <GridViewColumn.CellTemplate>
                    <DataTemplate>
                      <TextBox Style="{StaticResource CellBoxEdit}"
                               Text="{Binding QueueKey, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"/>
                    </DataTemplate>
                  </GridViewColumn.CellTemplate>
                </GridViewColumn>
                <GridViewColumn Header="Wert / VB-Ausdruck" Width="270">
                  <GridViewColumn.CellTemplate>
                    <DataTemplate>
                      <TextBox Style="{StaticResource CellBox}"
                               Text="{Binding Expression, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                               IsEnabled="{Binding ValueEditable}"/>
                    </DataTemplate>
                  </GridViewColumn.CellTemplate>
                </GridViewColumn>
                <GridViewColumn Header="Variable" Width="180" DisplayMemberBinding="{Binding VarName}"/>
              </GridView>
            </ListView.View>
          </ListView>

          <WrapPanel Grid.Row="2" Orientation="Horizontal">
            <Button x:Name="btnFieldUp"     Content="Nach oben"   MinWidth="112" Margin="0,0,6,0"/>
            <Button x:Name="btnFieldDown"   Content="Nach unten"  MinWidth="112"/>
            <Button x:Name="btnFieldAll"    Content="Alle auswaehlen" MinWidth="130"/>
            <Button x:Name="btnFieldNone"   Content="Alle abwaehlen"  MinWidth="130"/>
            <Border Width="18"/>
            <Button x:Name="btnFieldAddTxt" Content="+ Feld (fester Text)" MinWidth="160"/>
            <Button x:Name="btnFieldAddExp" Content="+ Feld (VB-Ausdruck)" MinWidth="160"/>
            <Button x:Name="btnFieldDel"    Content="Feld entfernen"      MinWidth="140"/>
          </WrapPanel>
        </Grid>
      </TabItem>

      <!-- ================= 3. Referenz ================= -->
      <TabItem Header="3 · Queue-Referenz">
        <Grid Margin="26,18,26,18">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <StackPanel Grid.Row="0">
            <TextBlock Style="{StaticResource Section}" Text="Reference des Queue-Items zusammenbauen"/>
            <TextBlock Style="{StaticResource Hint}"
                       Text="Links die verfuegbaren Bausteine (aktive Felder, Sonderwerte, eigener Text), rechts die Reihenfolge. Der Orchestrator erlaubt maximal 128 Zeichen pro Reference."/>
          </StackPanel>

          <Grid Grid.Row="1" Margin="0,4,0,10">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <DockPanel Grid.Column="0">
              <TextBlock DockPanel.Dock="Top" Style="{StaticResource Caption}" Text="VERFUEGBARE BAUSTEINE"/>
              <ListBox x:Name="lstRefAvailable" DisplayMemberPath="Display"/>
            </DockPanel>

            <StackPanel Grid.Column="1" VerticalAlignment="Center" Margin="14,0,14,0">
              <Button x:Name="btnRefAdd"    Content="Hinzufuegen" Style="{StaticResource BtnPrimary}" MinWidth="140" Margin="0,0,0,8"/>
              <Button x:Name="btnRefRemove" Content="Entfernen"   MinWidth="140" Margin="0,0,0,8"/>
              <Button x:Name="btnRefUp"     Content="Nach oben"   MinWidth="140" Margin="0,0,0,8"/>
              <Button x:Name="btnRefDown"   Content="Nach unten"  MinWidth="140"/>
            </StackPanel>

            <DockPanel Grid.Column="2">
              <TextBlock DockPanel.Dock="Top" Style="{StaticResource Caption}" Text="AUFBAU DER REFERENCE (VON OBEN NACH UNTEN)"/>
              <ListBox x:Name="lstRefSelected" DisplayMemberPath="Display"/>
            </DockPanel>
          </Grid>

          <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock Grid.Column="0" Text="Eigener Text:" Margin="0,0,10,0"/>
            <TextBox   Grid.Column="1" x:Name="txtRefCustom"/>
            <Button    Grid.Column="2" x:Name="btnRefAddText" Content="Text einfuegen" MinWidth="140"/>
            <CheckBox  Grid.Column="3" x:Name="chkRefTrim"     Content="Feldwerte trimmen" IsChecked="True" Margin="18,0,0,0"/>
            <CheckBox  Grid.Column="4" x:Name="chkRefTruncate" Content="auf 128 Zeichen kuerzen" IsChecked="True" Margin="18,0,0,0"/>
          </Grid>

          <StackPanel Grid.Row="3" Margin="0,14,0,0">
            <TextBlock Style="{StaticResource Caption}" Text="ERZEUGTER VB-AUSDRUCK"/>
            <TextBox x:Name="txtRefPreview" Style="{StaticResource MonoBox}" Height="72" TextWrapping="Wrap"
                     IsTabStop="False" Focusable="False"/>
          </StackPanel>
        </Grid>
      </TabItem>

      <!-- ================= 4. Konfiguration ================= -->
      <TabItem Header="4 · Konfiguration DEV/PRD">
        <Grid Margin="26,18,26,18">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <StackPanel Grid.Row="0">
            <TextBlock Style="{StaticResource Section}" Text="Data\Config_DEV.xlsx  ·  Data\Config_PRD.xlsx  (Sheet 'Settings')"/>
            <TextBlock Style="{StaticResource Hint}"
                       Text="Diese Name/Value-Paare landen im konfiguration_dict. SourceDataPath = Ordner mit den Input-Dateien, SourceSheetName = Tabellenblatt darin, QueueFolder = Orchestrator-Ordner, QueueName = Queue. Eigene Zeilen koennen ergaenzt werden."/>
          </StackPanel>

          <ListView Grid.Row="1" x:Name="lvConfig" Margin="0,4,0,10">
            <ListView.View>
              <GridView>
                <GridViewColumn Header="Name (Schluessel)" Width="250">
                  <GridViewColumn.CellTemplate>
                    <DataTemplate>
                      <TextBox Style="{StaticResource CellBox}"
                               Text="{Binding Name, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                               IsReadOnly="{Binding IsLocked}"/>
                    </DataTemplate>
                  </GridViewColumn.CellTemplate>
                </GridViewColumn>
                <GridViewColumn Header="Wert DEV  (SZT1 / SZT2)" Width="440">
                  <GridViewColumn.CellTemplate>
                    <DataTemplate>
                      <TextBox Style="{StaticResource CellBoxDev}"
                               Text="{Binding DevValue, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"/>
                    </DataTemplate>
                  </GridViewColumn.CellTemplate>
                </GridViewColumn>
                <GridViewColumn Header="Wert PRD  (SZP)" Width="440">
                  <GridViewColumn.CellTemplate>
                    <DataTemplate>
                      <TextBox Style="{StaticResource CellBoxPrd}"
                               Text="{Binding PrdValue, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"/>
                    </DataTemplate>
                  </GridViewColumn.CellTemplate>
                </GridViewColumn>
              </GridView>
            </ListView.View>
          </ListView>

          <WrapPanel Grid.Row="2">
            <Button x:Name="btnCfgAdd" Content="+ Zeile hinzufuegen" MinWidth="160" Margin="0,0,6,0"/>
            <Button x:Name="btnCfgDel" Content="Zeile entfernen"     MinWidth="160"/>
            <Button x:Name="btnCfgFromTemplate" Content="Werte aus Template uebernehmen" MinWidth="250"/>
          </WrapPanel>

          <Grid Grid.Row="3" Margin="0,16,0,0">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="200"/>
              <ColumnDefinition Width="230"/>
              <ColumnDefinition Width="200"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <TextBlock Grid.Row="0" Grid.Column="0" Text="Dateifilter Input-Ordner"/>
            <TextBox   Grid.Row="0" Grid.Column="1" x:Name="txtFileFilter" Text="*.xlsx"/>
            <TextBlock Grid.Row="0" Grid.Column="2" Text="Queue-Item Prioritaet"/>
            <ComboBox  Grid.Row="0" Grid.Column="3" x:Name="cboPriority" HorizontalAlignment="Left" Width="170">
              <ComboBoxItem Content="Low"/>
              <ComboBoxItem Content="Normal" IsSelected="True"/>
              <ComboBoxItem Content="High"/>
            </ComboBox>

            <TextBlock Grid.Row="1" Grid.Column="0" Text="Schluessel Queue-Name"/>
            <TextBox   Grid.Row="1" Grid.Column="1" x:Name="txtQueueNameKey" Text="QueueName"/>
            <TextBlock Grid.Row="1" Grid.Column="2" Text="Schluessel Queue-Folder"/>
            <TextBox   Grid.Row="1" Grid.Column="3" x:Name="txtQueueFolderKey" Text="QueueFolder"
                       HorizontalAlignment="Left" Width="170"/>
          </Grid>
        </Grid>
      </TabItem>

      <!-- ================= 5. Bestaetigung ================= -->
      <TabItem Header="5 · Bestaetigung">
        <Grid Margin="26,18,26,18">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"  MinHeight="200"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"  MinHeight="240"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <Grid Grid.Row="0">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0">
              <TextBlock Style="{StaticResource Section}" Text="Bitte pruefen und bestaetigen"/>
              <TextBlock Style="{StaticResource Hint}"
                         Text="Nach 'Dispatcher erstellen' wird das Template kopiert und im Ziel-Ordner angepasst. Das Template selbst bleibt unveraendert."/>
            </StackPanel>
            <Button Grid.Column="1" x:Name="btnRefreshSummary" Content="Zusammenfassung aktualisieren"
                    MinWidth="230" VerticalAlignment="Center"/>
          </Grid>

          <TextBox Grid.Row="1" x:Name="txtSummary" Style="{StaticResource MonoBox}"/>

          <GridSplitter Grid.Row="2" Height="8" HorizontalAlignment="Stretch"
                        VerticalAlignment="Center" Background="Transparent"
                        ResizeBehavior="PreviousAndNext" ResizeDirection="Rows"
                        ShowsPreview="True" Cursor="SizeNS"/>

          <DockPanel Grid.Row="3">
            <Grid DockPanel.Dock="Top">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <TextBlock Grid.Column="0" Style="{StaticResource Caption}" Text="PROTOKOLL"/>
              <TextBlock Grid.Column="1" Style="{StaticResource Caption}" HorizontalAlignment="Right"
                         Text="TRENNLINIE OBEN ZIEHBAR"/>
            </Grid>
            <TextBox x:Name="txtLog" Style="{StaticResource MonoBox}" Background="{StaticResource Weiss}"/>
          </DockPanel>

          <CheckBox Grid.Row="4" x:Name="chkConfirm" Margin="0,14,0,0" FontWeight="Bold"
                    Content="Ich habe die Angaben geprueft und moechte den Dispatcher jetzt erzeugen."/>
        </Grid>
      </TabItem>

    </TabControl>
  </DockPanel>
</Window>
'@

$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# --- Controls einsammeln ---
$ctl = @{}
foreach ($name in @(
    'txtStatus','btnBack','btnNext','btnCreate','btnClose','tabMain',
    'txtTemplate','btnTemplate','txtOutput','btnOutput','txtProjectName','txtProjectDesc',
    'txtSourceExcel','btnSourceExcel','cboSheet','btnReadColumns','txtDataFolder','txtConfigSheet','cboEnv',
    'lvFields','btnFieldUp','btnFieldDown','btnFieldAll','btnFieldNone','btnFieldAddTxt','btnFieldAddExp','btnFieldDel',
    'lstRefAvailable','lstRefSelected','btnRefAdd','btnRefRemove','btnRefUp','btnRefDown',
    'txtRefCustom','btnRefAddText','chkRefTrim','chkRefTruncate','txtRefPreview',
    'lvConfig','btnCfgAdd','btnCfgDel','btnCfgFromTemplate','txtFileFilter','cboPriority',
    'txtQueueNameKey','txtQueueFolderKey',
    'txtSummary','chkConfirm','txtLog','btnRefreshSummary'
)) { $ctl[$name] = $window.FindName($name) }

# --- Collections ---
$script:Fields    = New-Object 'System.Collections.ObjectModel.ObservableCollection[UiPathGen.FieldItem]'
$script:ConfigRow = New-Object 'System.Collections.ObjectModel.ObservableCollection[UiPathGen.ConfigItem]'
$script:RefAvail  = New-Object 'System.Collections.ObjectModel.ObservableCollection[UiPathGen.RefBlock]'
$script:RefSel    = New-Object 'System.Collections.ObjectModel.ObservableCollection[UiPathGen.RefBlock]'
$script:Columns   = @()

$ctl.lvFields.ItemsSource        = $script:Fields
$ctl.lvConfig.ItemsSource        = $script:ConfigRow
$ctl.lstRefAvailable.ItemsSource = $script:RefAvail
$ctl.lstRefSelected.ItemsSource  = $script:RefSel

function Set-Status { param([string]$Text) $ctl.txtStatus.Text = $Text }

function Get-ComboText {
    <# Liefert den Text der Auswahl - egal ob die Items Strings oder ComboBoxItems sind. #>
    param($Combo)
    $sel = $Combo.SelectedItem
    if ($null -eq $sel) { return [string]$Combo.Text }
    if ($sel -is [System.Windows.Controls.ComboBoxItem]) { return [string]$sel.Content }
    return [string]$sel
}

function Show-Info  { param([string]$m) [System.Windows.MessageBox]::Show($window, $m, 'Hinweis',  'OK', 'Information') | Out-Null }
function Show-Warn  { param([string]$m) [System.Windows.MessageBox]::Show($window, $m, 'Achtung',  'OK', 'Warning')     | Out-Null }
function Show-Error { param([string]$m) [System.Windows.MessageBox]::Show($window, $m, 'Fehler',   'OK', 'Error')       | Out-Null }

# ---------------------------------------------------------------------------
#  Felder: Pflichteintraege
# ---------------------------------------------------------------------------
function Reset-MandatoryField {
    $script:Fields.Clear()

    $f1 = New-Object UiPathGen.FieldItem
    $f1.Kind = 'Auto'; $f1.ColumnName = '(automatisch: Dateiname der Quelldatei)'
    $f1.QueueKey = 'filename'; $f1.VarName = '@@FILENAME@@'; $f1.Expression = 'in_dokumentname'
    $f1.IsLocked = $true; $f1.Selected = $true
    $script:Fields.Add($f1)

    $f2 = New-Object UiPathGen.FieldItem
    $f2.Kind = 'Auto'; $f2.ColumnName = '(automatisch: Erstellungszeitpunkt)'
    $f2.QueueKey = 'CreateTS'; $f2.VarName = '@@CREATETS@@'; $f2.Expression = 'DateTime.Now'
    $f2.IsLocked = $true; $f2.Selected = $true
    $script:Fields.Add($f2)
}

function New-UsedNameSet {
    <#
      Namen, die fuer eine neue Feld-Variable nicht mehr frei sind: die bereits
      vergebenen VarNames plus die festen Infrastruktur-Variablen. Ohne letztere
      erzeugte eine Excel-Spalte 'Reference' ein zweites <Variable Name="Reference" />
      und Studio verweigerte das Oeffnen des Workflows.
    #>
    $used = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($k in $script:KeepVars) { [void]$used.Add($k.ToLowerInvariant()) }
    foreach ($f in $script:Fields) {
        if ($f.VarName -and $f.VarName -notlike '@@*') { [void]$used.Add($f.VarName.ToLowerInvariant()) }
    }
    # Komma: sonst entrollt PowerShell das (aufzaehlbare) HashSet.
    return ,$used
}

function Add-ColumnField {
    param([string[]]$Columns)

    $used = New-UsedNameSet
    foreach ($c in $Columns) {
        # Leere Kopfzellen ergeben keine Spalte. Frueher wurde daraus ein
        # erfundenes 'SpalteN', das die DataTable zur Laufzeit nicht kennt.
        if ([string]::IsNullOrWhiteSpace($c)) { continue }
        $item = New-Object UiPathGen.FieldItem
        $item.Kind       = 'Spalte'
        $item.ColumnName = $c
        $item.QueueKey   = $c
        $item.VarName    = ConvertTo-VbIdentifier -Name $c -Used $used
        $item.Selected   = $true
        $script:Fields.Add($item)
    }
}

function Sync-RefSelected {
    <#
      Entfernt Reference-Bausteine, deren Feld nicht mehr aktiv ist. Ohne das
      erzeugte der Generator ein [Kundennummer & ...], obwohl die Variable
      Kundennummer gar nicht mehr deklariert wird -> "is not declared" in Studio.
    #>
    $activeVars = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($f in (Get-ActiveField $script:Fields)) { if ($f.VarName) { [void]$activeVars.Add($f.VarName) } }
    foreach ($b in @($script:RefSel)) {
        if ($b.Kind -eq 'Feld' -and -not $activeVars.Contains($b.Expression)) { [void]$script:RefSel.Remove($b) }
    }
    Update-RefPreview
}

function Update-RefAvailable {
    $script:RefAvail.Clear()

    foreach ($f in $script:Fields) {
        if (-not $f.Selected) { continue }
        $b = New-Object UiPathGen.RefBlock
        if ($f.Kind -eq 'Auto') {
            $b.Kind = 'Sonder'
            $b.Display = "[$($f.QueueKey)]"
            if ($f.VarName -eq '@@FILENAME@@') { $b.Expression = 'in_dokumentname' }
            else { $b.Expression = 'DateTime.Now.ToString("yyyyMMddHHmmss")' }
        }
        else {
            $b.Kind = 'Feld'
            $b.Display = "[$($f.QueueKey)]  ($($f.VarName))"
            $b.Expression = $f.VarName
        }
        $script:RefAvail.Add($b)
    }

    foreach ($s in @(
        @{ D = '{Zeitstempel yyyyMMddHHmmss}'; E = 'DateTime.Now.ToString("yyyyMMddHHmmss")' },
        @{ D = '{Datum yyyy-MM-dd}';           E = 'DateTime.Now.ToString("yyyy-MM-dd")' },
        @{ D = '{Neue GUID}';                  E = 'Guid.NewGuid.ToString("N")' },
        @{ D = '{Trennzeichen _}';             E = '"_"' },
        @{ D = '{Trennzeichen -}';             E = '"-"' },
        @{ D = '{Trennzeichen |}';             E = '"|"' }
    )) {
        $b = New-Object UiPathGen.RefBlock
        $b.Kind = 'Sonder'; $b.Display = $s.D; $b.Expression = $s.E
        $script:RefAvail.Add($b)
    }
}

function Update-RefPreview {
    $expr = New-ReferenceExpression -Blocks @($script:RefSel) -TrimParts ([bool]$ctl.chkRefTrim.IsChecked)
    if (-not $expr) {
        $ctl.txtRefPreview.Text = '(noch kein Baustein gewaehlt - Reference waere leer)'
    }
    else {
        $t = "[$expr]"
        if ($ctl.chkRefTruncate.IsChecked) {
            $t += "`r`n" + '  danach:  [If(Reference.Length > 128, Reference.Substring(0, 128), Reference)]'
        }
        $ctl.txtRefPreview.Text = $t
    }
}

function Reset-ConfigRow {
    param([hashtable]$Dev, [hashtable]$Prd)

    $script:ConfigRow.Clear()
    $defaults = @(
        @{ N = 'SourceDataPath';  D = 'C:\RPA\Input\';  P = 'C:\RPA\Input\';  L = $true },
        @{ N = 'SourceSheetName'; D = 'Sheet1';         P = 'Sheet1';         L = $true },
        @{ N = 'QueueFolder';     D = 'Shared';         P = 'Shared';         L = $true },
        @{ N = 'QueueName';       D = '';               P = '';               L = $true }
    )
    foreach ($d in $defaults) {
        $c = New-Object UiPathGen.ConfigItem
        $c.Name = $d.N
        $c.DevValue = $d.D
        $c.PrdValue = $d.P
        $c.IsLocked = $d.L
        if ($Dev -and $Dev.ContainsKey($d.N)) { $c.DevValue = [string]$Dev[$d.N] }
        if ($Prd -and $Prd.ContainsKey($d.N)) { $c.PrdValue = [string]$Prd[$d.N] }
        $script:ConfigRow.Add($c)
    }

    # zusaetzliche Schluessel aus dem Template uebernehmen
    $known = @('SourceDataPath', 'SourceSheetName', 'QueueFolder', 'QueueName')
    $extra = New-Object 'System.Collections.Generic.List[string]'
    foreach ($h in @($Dev, $Prd)) {
        if (-not $h) { continue }
        foreach ($k in $h.Keys) { if ($known -notcontains $k -and -not $extra.Contains($k)) { $extra.Add($k) } }
    }
    foreach ($k in $extra) {
        $c = New-Object UiPathGen.ConfigItem
        $c.Name = $k
        if ($Dev -and $Dev.ContainsKey($k)) { $c.DevValue = [string]$Dev[$k] }
        if ($Prd -and $Prd.ContainsKey($k)) { $c.PrdValue = [string]$Prd[$k] }
        $script:ConfigRow.Add($c)
    }
}

function Get-ConfigValue {
    param([string]$Name, [string]$Env = 'DEV')
    foreach ($c in $script:ConfigRow) {
        if ($c.Name -eq $Name) { if ($Env -eq 'DEV') { return $c.DevValue } else { return $c.PrdValue } }
    }
    return ''
}

function Set-ConfigValue {
    param([string]$Name, [string]$Value, [switch]$OnlyIfEmpty)
    foreach ($c in $script:ConfigRow) {
        if ($c.Name -eq $Name) {
            if (-not $OnlyIfEmpty -or [string]::IsNullOrWhiteSpace($c.DevValue)) { $c.DevValue = $Value }
            if (-not $OnlyIfEmpty -or [string]::IsNullOrWhiteSpace($c.PrdValue)) { $c.PrdValue = $Value }
            return
        }
    }
}

# ---------------------------------------------------------------------------
#  Zusammenfassung
# ---------------------------------------------------------------------------
function Format-Summary {
    $sb = New-Object System.Text.StringBuilder
    $active = Get-ActiveField $script:Fields

    [void]$sb.AppendLine('PROJEKT')
    [void]$sb.AppendLine('------------------------------------------------------------------')
    [void]$sb.AppendLine("  Template          : $($ctl.txtTemplate.Text)")
    [void]$sb.AppendLine("  Ziel-Ordner       : $($ctl.txtOutput.Text)")
    [void]$sb.AppendLine("  Projektname       : $($ctl.txtProjectName.Text)")
    [void]$sb.AppendLine("  Beschreibung      : $($ctl.txtProjectDesc.Text)")
    [void]$sb.AppendLine("  Default Umgebung  : $(Get-ComboText $ctl.cboEnv)")
    [void]$sb.AppendLine("  Config            : $($ctl.txtDataFolder.Text)\Config_DEV.xlsx / Config_PRD.xlsx  (Sheet '$($ctl.txtConfigSheet.Text)')")
    [void]$sb.AppendLine('')

    [void]$sb.AppendLine('QUELL-EXCEL (Analyse)')
    [void]$sb.AppendLine('------------------------------------------------------------------')
    [void]$sb.AppendLine("  Datei             : $($ctl.txtSourceExcel.Text)")
    [void]$sb.AppendLine("  Sheet             : $(Get-ComboText $ctl.cboSheet)")
    [void]$sb.AppendLine("  Erkannte Spalten  : $($script:Columns.Count)")
    [void]$sb.AppendLine('')

    [void]$sb.AppendLine("QUEUE-ITEM INHALT  ($($active.Count) Eintraege, in dieser Reihenfolge)")
    [void]$sb.AppendLine('------------------------------------------------------------------')
    $n = 0
    foreach ($f in $active) {
        $n++
        $src = switch ($f.Kind) {
            'Auto'     { $f.Expression }
            'Spalte'   { "Excel-Spalte `"$($f.ColumnName)`"" }
            'Text'     { "fester Text `"$($f.Expression)`"" }
            'Ausdruck' { "VB-Ausdruck $($f.Expression)" }
        }
        [void]$sb.AppendLine(("  {0,2}. {1,-32} <- {2}" -f $n, $f.QueueKey, $src))
    }
    $inactive = @($script:Fields | Where-Object { -not $_.Selected })
    if ($inactive.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("  Nicht uebernommen : " + (($inactive | ForEach-Object { $_.ColumnName }) -join ', '))
    }
    [void]$sb.AppendLine('')

    [void]$sb.AppendLine('QUEUE-ITEM REFERENCE')
    [void]$sb.AppendLine('------------------------------------------------------------------')
    $expr = New-ReferenceExpression -Blocks @($script:RefSel) -TrimParts ([bool]$ctl.chkRefTrim.IsChecked)
    [void]$sb.AppendLine("  Bausteine         : " + (($script:RefSel | ForEach-Object { $_.Display }) -join '  +  '))
    [void]$sb.AppendLine("  Ausdruck          : [$expr]")
    $cut = 'nein'; if ($ctl.chkRefTruncate.IsChecked) { $cut = 'ja (128 Zeichen)' }
    [void]$sb.AppendLine("  Kuerzen           : $cut")
    [void]$sb.AppendLine("  Prioritaet        : $(Get-ComboText $ctl.cboPriority)")
    [void]$sb.AppendLine('')

    [void]$sb.AppendLine('KONFIGURATION')
    [void]$sb.AppendLine('------------------------------------------------------------------')
    [void]$sb.AppendLine(("  {0,-22} {1,-32} {2}" -f 'Name', 'DEV', 'PRD'))
    foreach ($c in $script:ConfigRow) {
        if ([string]::IsNullOrWhiteSpace($c.Name)) { continue }
        [void]$sb.AppendLine(("  {0,-22} {1,-32} {2}" -f $c.Name, $c.DevValue, $c.PrdValue))
    }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("  Dateifilter       : $($ctl.txtFileFilter.Text)")
    [void]$sb.AppendLine("  Queue-Name aus    : konfiguration_dict(""$($ctl.txtQueueNameKey.Text)"")")
    [void]$sb.AppendLine("  Queue-Folder aus  : konfiguration_dict(""$($ctl.txtQueueFolderKey.Text)"")")
    [void]$sb.AppendLine('')

    [void]$sb.AppendLine('GENERIERTES qItemInfo-DICTIONARY (PopulateQueue.xaml)')
    [void]$sb.AppendLine('------------------------------------------------------------------')
    foreach ($line in (New-QItemDictionary -Fields $active -Indent '    ') -split "`r`n") {
        [void]$sb.AppendLine("  $line")
    }
    [void]$sb.AppendLine('')

    [void]$sb.AppendLine('WAS ANGEPASST WIRD')
    [void]$sb.AppendLine('------------------------------------------------------------------')
    [void]$sb.AppendLine('  project.json         Name, Beschreibung, neue Projekt-/EntryPoint-GUIDs')
    [void]$sb.AppendLine('  Main.xaml            Default In_Umgebung, Invoke-Pfade, Config-Datei/-Sheet')
    [void]$sb.AppendLine('  ReadSourceData.xaml  Dateifilter, Zuweisung out_dokumentname')
    [void]$sb.AppendLine('  PopulateQueue.xaml   Variablen, MultipleAssign, qItemInfo, Reference,')
    [void]$sb.AppendLine('                       AddQueueItem (Queue/Folder/Prioritaet), Logmeldungen')
    [void]$sb.AppendLine('  Config_DEV/PRD.xlsx  Settings-Sheet neu geschrieben')

    return $sb.ToString()
}

# ---------------------------------------------------------------------------
#  Validierung pro Schritt
# ---------------------------------------------------------------------------
function Test-Step {
    param([int]$Index)

    switch ($Index) {
        0 {
            if ([string]::IsNullOrWhiteSpace($ctl.txtTemplate.Text) -or -not (Test-Path -LiteralPath $ctl.txtTemplate.Text)) {
                Show-Warn 'Bitte einen gueltigen Template-Ordner waehlen.'; return $false
            }
            # -ErrorAction SilentlyContinue verschluckte hier Zugriffsfehler und
            # meldete sie als "keine PopulateQueue.xaml gefunden" - man sucht dann
            # den Fehler an der falschen Stelle.
            try {
                $found = @(Get-ChildItem -LiteralPath $ctl.txtTemplate.Text -Recurse -Filter 'PopulateQueue.xaml' -ErrorAction Stop)
            }
            catch {
                Show-Warn "Der Template-Ordner konnte nicht gelesen werden:`r`n$($_.Exception.Message)"; return $false
            }
            if ($found.Count -eq 0) {
                Show-Warn 'Im Template-Ordner wurde keine PopulateQueue.xaml gefunden.'; return $false
            }
            if ([string]::IsNullOrWhiteSpace($ctl.txtOutput.Text)) { Show-Warn 'Bitte einen Ziel-Ordner angeben.'; return $false }
            if ([string]::IsNullOrWhiteSpace($ctl.txtProjectName.Text)) { Show-Warn 'Bitte einen Projektnamen angeben.'; return $false }
            if ($script:Columns.Count -eq 0) {
                Show-Warn 'Bitte Quell-Excel und Sheet waehlen und "Spalten einlesen" ausfuehren.'; return $false
            }
            # Der Config-Ordner wird als Verzeichnis angelegt. Unzulaessige Zeichen
            # liessen die Generierung erst mitten im Lauf scheitern.
            $dataName = $ctl.txtDataFolder.Text
            if (-not [string]::IsNullOrWhiteSpace($dataName) -and
                $dataName.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0) {
                Show-Warn 'Der Config-Ordner enthaelt unzulaessige Zeichen fuer einen Ordnernamen.'; return $false
            }
            return $true
        }
        1 {
            $active = Get-ActiveField $script:Fields
            if ($active.Count -lt 3) {
                Show-Warn 'Bitte mindestens ein Feld zusaetzlich zu filename und CreateTS auswaehlen.'; return $false
            }
            $seen = New-Object 'System.Collections.Generic.HashSet[string]'
            foreach ($f in $active) {
                if (-not $seen.Add($f.QueueKey.ToLowerInvariant())) {
                    Show-Warn "Der Queue-Schluessel '$($f.QueueKey)' kommt mehrfach vor. Schluessel muessen eindeutig sein."; return $false
                }
            }
            return $true
        }
        2 {
            if ($script:RefSel.Count -eq 0) { Show-Warn 'Bitte mindestens einen Baustein fuer die Reference waehlen.'; return $false }
            # Das Aktiv-Haekchen auf Maske 2 hat keinen Handler, der die Bausteine
            # nachzieht. Ohne diese Pruefung entstand ein Verweis auf eine
            # Variable, die im generierten Workflow nicht mehr deklariert ist.
            $activeVars = New-Object 'System.Collections.Generic.HashSet[string]'
            foreach ($f in (Get-ActiveField $script:Fields)) { if ($f.VarName) { [void]$activeVars.Add($f.VarName) } }
            foreach ($b in $script:RefSel) {
                if ($b.Kind -eq 'Feld' -and -not $activeVars.Contains($b.Expression)) {
                    Show-Warn "Der Reference-Baustein '$($b.Display)' zeigt auf ein Feld, das nicht mehr aktiv ist.`r`nBitte den Baustein entfernen oder das Feld auf Maske 2 wieder aktivieren."
                    return $false
                }
            }
            return $true
        }
        3 {
            foreach ($k in @('SourceDataPath', 'SourceSheetName', 'QueueName')) {
                if ([string]::IsNullOrWhiteSpace((Get-ConfigValue -Name $k -Env 'DEV'))) {
                    Show-Warn "Bitte '$k' fuer DEV ausfuellen."; return $false
                }
            }
            return $true
        }
    }
    return $true
}

# ---------------------------------------------------------------------------
#  Events
# ---------------------------------------------------------------------------

$ctl.btnTemplate.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'Ordner des Dispatcher-Templates waehlen'
    if ($ctl.txtTemplate.Text -and (Test-Path -LiteralPath $ctl.txtTemplate.Text)) { $dlg.SelectedPath = $ctl.txtTemplate.Text }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $ctl.txtTemplate.Text = $dlg.SelectedPath
        if ([string]::IsNullOrWhiteSpace($ctl.txtProjectName.Text)) {
            $ctl.txtProjectName.Text = (Split-Path -Leaf $dlg.SelectedPath)
        }
        & $script:LoadTemplateConfig
    }
})

$script:LoadTemplateConfig = {
    try {
        $dataName = $ctl.txtDataFolder.Text
        if ([string]::IsNullOrWhiteSpace($dataName)) { $dataName = 'Data' }
        $sheet = $ctl.txtConfigSheet.Text
        if ([string]::IsNullOrWhiteSpace($sheet)) { $sheet = 'Settings' }

        $devPath = Join-Path (Join-Path $ctl.txtTemplate.Text $dataName) 'Config_DEV.xlsx'
        $prdPath = Join-Path (Join-Path $ctl.txtTemplate.Text $dataName) 'Config_PRD.xlsx'
        $dev = $null; $prd = $null
        if (Test-Path -LiteralPath $devPath) { $dev = Get-XlsxNameValue -Path $devPath -SheetName $sheet }
        if (Test-Path -LiteralPath $prdPath) { $prd = Get-XlsxNameValue -Path $prdPath -SheetName $sheet }

        $devH = @{}; if ($dev) { foreach ($k in $dev.Keys) { $devH[$k] = $dev[$k] } }
        $prdH = @{}; if ($prd) { foreach ($k in $prd.Keys) { $prdH[$k] = $prd[$k] } }
        Reset-ConfigRow -Dev $devH -Prd $prdH
        Set-Status "Konfiguration aus Template gelesen ($($devH.Count) DEV- / $($prdH.Count) PRD-Eintraege)."
    }
    catch {
        Reset-ConfigRow -Dev @{} -Prd @{}
        Set-Status "Template-Konfiguration konnte nicht gelesen werden: $($_.Exception.Message)"
    }
}

$ctl.btnOutput.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'Ziel-Ordner fuer den neuen Dispatcher waehlen'
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $ctl.txtOutput.Text = $dlg.SelectedPath }
})

$ctl.btnSourceExcel.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Excel-Arbeitsmappe (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|Alle Dateien (*.*)|*.*'
    $dlg.Title  = 'Quell-Excel waehlen'
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $ctl.txtSourceExcel.Text = $dlg.FileName
        try {
            $sheets = Get-XlsxSheet -Path $dlg.FileName
            $ctl.cboSheet.Items.Clear()
            foreach ($s in $sheets) { [void]$ctl.cboSheet.Items.Add($s.Name) }
            if ($ctl.cboSheet.Items.Count -gt 0) { $ctl.cboSheet.SelectedIndex = 0 }
            Set-Status "$($sheets.Count) Tabellenblatt/-blaetter gefunden."

            # Vorbelegung der Konfiguration aus der Quelldatei
            Set-ConfigValue -Name 'SourceDataPath'  -Value ((Split-Path -Parent $dlg.FileName) + '\')
            if ($ctl.cboSheet.SelectedItem) { Set-ConfigValue -Name 'SourceSheetName' -Value ([string]$ctl.cboSheet.SelectedItem) }
        }
        catch { Show-Error "Excel konnte nicht gelesen werden:`r`n$($_.Exception.Message)" }
    }
})

$ctl.cboSheet.Add_SelectionChanged({
    if ($ctl.cboSheet.SelectedItem) { Set-ConfigValue -Name 'SourceSheetName' -Value ([string]$ctl.cboSheet.SelectedItem) }
})

$ctl.btnReadColumns.Add_Click({
    if ([string]::IsNullOrWhiteSpace($ctl.txtSourceExcel.Text) -or -not $ctl.cboSheet.SelectedItem) {
        Show-Warn 'Bitte zuerst Excel-Datei und Tabellenblatt waehlen.'; return
    }
    # Erneutes Einlesen verwirft alle Felder und die ganze Reference. Vorher
    # geschah das kommentarlos, mitten in einer halbfertigen Konfiguration.
    if ($script:Fields.Count -gt 2 -or $script:RefSel.Count -gt 0) {
        $r = [System.Windows.MessageBox]::Show($window,
            "Die bereits konfigurierten Felder und die zusammengestellte Reference gehen dabei verloren.`r`n`r`nSpalten wirklich neu einlesen?",
            'Konfiguration verwerfen?', 'YesNo', 'Warning')
        if ($r -ne [System.Windows.MessageBoxResult]::Yes) { return }
    }
    try {
        $cols = Get-XlsxHeader -Path $ctl.txtSourceExcel.Text -SheetName ([string]$ctl.cboSheet.SelectedItem)
        $named = @($cols | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($named.Count -eq 0) { Show-Warn 'In der ersten Zeile des Sheets wurden keine Spaltenueberschriften gefunden.'; return }
        $script:Columns = @($cols)
        Reset-MandatoryField
        Add-ColumnField -Columns $script:Columns
        Update-RefAvailable
        $script:RefSel.Clear()
        Update-RefPreview
        $blank = $cols.Count - $named.Count
        $hint = ''
        if ($blank -gt 0) { $hint = " ($blank Spalte(n) ohne Kopfzeile uebersprungen)" }
        Set-Status "$($named.Count) Spalten eingelesen$($hint): $($named -join ', ')"
        Show-Info "$($named.Count) Spalten eingelesen$hint. Weiter zu Schritt 2."
    }
    catch { Show-Error "Spalten konnten nicht gelesen werden:`r`n$($_.Exception.Message)" }
})

# --- Felder ---
$moveField = {
    param([int]$Delta)
    $sel = $ctl.lvFields.SelectedItem
    if (-not $sel) { return }
    $i = $script:Fields.IndexOf($sel)
    $j = $i + $Delta
    if ($sel.IsLocked) { return }
    if ($j -lt 2 -or $j -ge $script:Fields.Count) { return }
    $script:Fields.Move($i, $j)
    $ctl.lvFields.SelectedItem = $sel
}
$ctl.btnFieldUp.Add_Click({   & $moveField -1 })
$ctl.btnFieldDown.Add_Click({ & $moveField  1 })

$ctl.btnFieldAll.Add_Click({  foreach ($f in $script:Fields) { $f.Selected = $true  }; Update-RefAvailable; Sync-RefSelected })
$ctl.btnFieldNone.Add_Click({ foreach ($f in $script:Fields) { if (-not $f.IsLocked) { $f.Selected = $false } }; Update-RefAvailable; Sync-RefSelected })

$addCustomField = {
    param([string]$Kind)
    $used = New-UsedNameSet
    $item = New-Object UiPathGen.FieldItem
    $item.Kind       = $Kind
    $item.ColumnName = "(eigenes Feld)"
    $item.QueueKey   = "NeuesFeld$($script:Fields.Count)"
    $item.VarName    = ConvertTo-VbIdentifier -Name $item.QueueKey -Used $used
    if ($Kind -eq 'Text') { $item.Expression = 'Wert eintragen' } else { $item.Expression = 'DateTime.Now.ToString("yyyy-MM-dd")' }
    $item.Selected   = $true
    $script:Fields.Add($item)
    $ctl.lvFields.SelectedItem = $item
    $ctl.lvFields.ScrollIntoView($item)
    Set-Status 'Feld hinzugefuegt - Schluessel und Wert direkt in der Tabelle anpassen.'
}
$ctl.btnFieldAddTxt.Add_Click({ & $addCustomField 'Text' })
$ctl.btnFieldAddExp.Add_Click({ & $addCustomField 'Ausdruck' })

$ctl.btnFieldDel.Add_Click({
    $sel = @($ctl.lvFields.SelectedItems)
    if ($sel.Count -eq 0) { Show-Warn 'Bitte zuerst eine Zeile markieren.'; return }
    foreach ($s in $sel) {
        if ($s.IsLocked) { continue }
        [void]$script:Fields.Remove($s)
    }
    Update-RefAvailable
    Sync-RefSelected
})

# --- Referenz ---
$ctl.btnRefAdd.Add_Click({
    $sel = $ctl.lstRefAvailable.SelectedItem
    if (-not $sel) { return }
    $b = New-Object UiPathGen.RefBlock
    $b.Kind = $sel.Kind; $b.Display = $sel.Display; $b.Expression = $sel.Expression
    $script:RefSel.Add($b)
    Update-RefPreview
})
$ctl.btnRefRemove.Add_Click({
    $sel = $ctl.lstRefSelected.SelectedItem
    if ($sel) { [void]$script:RefSel.Remove($sel); Update-RefPreview }
})
$moveRef = {
    param([int]$Delta)
    $sel = $ctl.lstRefSelected.SelectedItem
    if (-not $sel) { return }
    $i = $script:RefSel.IndexOf($sel); $j = $i + $Delta
    if ($j -lt 0 -or $j -ge $script:RefSel.Count) { return }
    $script:RefSel.Move($i, $j)
    $ctl.lstRefSelected.SelectedItem = $sel
    Update-RefPreview
}
$ctl.btnRefUp.Add_Click({   & $moveRef -1 })
$ctl.btnRefDown.Add_Click({ & $moveRef  1 })

$ctl.btnRefAddText.Add_Click({
    $t = $ctl.txtRefCustom.Text
    if ([string]::IsNullOrEmpty($t)) { Show-Warn 'Bitte einen Text eingeben.'; return }
    $b = New-Object UiPathGen.RefBlock
    $b.Kind = 'Text'; $b.Display = "`"$t`""; $b.Expression = (ConvertTo-VbLiteral $t)
    $script:RefSel.Add($b)
    $ctl.txtRefCustom.Text = ''
    Update-RefPreview
})
# Checked/Unchecked statt Click: Click feuert nicht, wenn IsChecked im Code
# gesetzt wird, die Vorschau zeigte dann einen veralteten Ausdruck.
$ctl.chkRefTrim.Add_Checked({ Update-RefPreview });     $ctl.chkRefTrim.Add_Unchecked({ Update-RefPreview })
$ctl.chkRefTruncate.Add_Checked({ Update-RefPreview }); $ctl.chkRefTruncate.Add_Unchecked({ Update-RefPreview })

# --- Konfiguration ---
$ctl.btnCfgAdd.Add_Click({
    $c = New-Object UiPathGen.ConfigItem
    $c.Name = 'NeuerSchluessel'
    $script:ConfigRow.Add($c)
    $ctl.lvConfig.SelectedItem = $c
    $ctl.lvConfig.ScrollIntoView($c)
})
$ctl.btnCfgDel.Add_Click({
    $sel = @($ctl.lvConfig.SelectedItems)
    foreach ($s in $sel) { if (-not $s.IsLocked) { [void]$script:ConfigRow.Remove($s) } }
})
$ctl.btnCfgFromTemplate.Add_Click({ & $script:LoadTemplateConfig })

# --- Navigation ---
function Update-Summary {
    <# Baut die Zusammenfassung neu auf. Fehler landen sichtbar im Feld statt lautlos zu verpuffen. #>
    try {
        $ctl.txtSummary.Text = Format-Summary
    }
    catch {
        $ctl.txtSummary.Text = "Die Zusammenfassung konnte nicht erstellt werden:" + [Environment]::NewLine +
                               $_.Exception.Message + [Environment]::NewLine + [Environment]::NewLine +
                               $_.ScriptStackTrace
    }
}

function Update-Navigation {
    $i = $ctl.tabMain.SelectedIndex
    if ($i -lt 0) { return }
    $last = $ctl.tabMain.Items.Count - 1
    $ctl.btnBack.IsEnabled = ($i -gt 0)
    $ctl.btnNext.IsEnabled = ($i -lt $last)
    if ($i -eq $last) {
        $ctl.btnCreate.Visibility = [System.Windows.Visibility]::Visible
        Update-Summary
    }
    else {
        $ctl.btnCreate.Visibility = [System.Windows.Visibility]::Collapsed
    }
}

# Kein Source-Filter mehr: Update-Navigation ist nebenwirkungsfrei und idempotent,
# dadurch bleibt die Zusammenfassung auch bei Aenderungen in den Listen aktuell.
$ctl.tabMain.Add_SelectionChanged({ Update-Navigation })

$ctl.btnNext.Add_Click({
    $i = $ctl.tabMain.SelectedIndex
    if (Test-Step $i) {
        if ($i -eq 1) { Update-RefAvailable }
        if ($i -lt ($ctl.tabMain.Items.Count - 1)) { $ctl.tabMain.SelectedIndex = $i + 1 }
    }
    Update-Navigation
})
$ctl.btnBack.Add_Click({
    if ($ctl.tabMain.SelectedIndex -gt 0) { $ctl.tabMain.SelectedIndex = $ctl.tabMain.SelectedIndex - 1 }
    Update-Navigation
})
$ctl.btnRefreshSummary.Add_Click({ Update-Summary; Set-Status 'Zusammenfassung aktualisiert.' })
$ctl.chkConfirm.Add_Click({ Update-Summary })
$ctl.btnClose.Add_Click({ $window.Close() })

# --- Erstellen ---
$ctl.btnCreate.Add_Click({
    # Alles ab hier laeuft in einem try. Frueher standen die Vorpruefungen davor;
    # mit $ErrorActionPreference='Stop' beendete etwa ein Rechtefehler auf einem
    # UNC-Zielpfad den Handler wortlos - der Klick sah einfach wirkungslos aus.
    $out = ''
    try {
        if (-not $ctl.chkConfirm.IsChecked) { Show-Warn 'Bitte die Bestaetigung ankreuzen.'; return }
        for ($i = 0; $i -le 3; $i++) {
            if (-not (Test-Step $i)) { $ctl.tabMain.SelectedIndex = $i; return }
        }

        $out = $ctl.txtOutput.Text
        if ((Test-Path -LiteralPath $out) -and (Get-ChildItem -LiteralPath $out -Force | Select-Object -First 1)) {
            # Die alte Warnung sprach nur von "Dateien mit gleichem Namen". Was
            # wirklich passiert, ist deutlich mehr - das gehoert vor die Entscheidung.
            $r = [System.Windows.MessageBox]::Show($window,
                ("Der Ziel-Ordner ist nicht leer:`r`n$out`r`n`r`n" +
                 "Main.xaml, project.json und die Utility-Workflows werden durch die Template-Versionen ersetzt. " +
                 "Unterordner wie Utility_Files und Data werden vorher vollstaendig geloescht, " +
                 "Config_DEV.xlsx und Config_PRD.xlsx werden neu geschrieben, und das Projekt bekommt eine neue projectId.`r`n`r`n" +
                 "Eigene Aenderungen in diesem Ordner gehen dabei verloren.`r`n`r`nFortfahren?"),
                'Ziel-Ordner nicht leer', 'YesNo', 'Warning')
            if ($r -ne [System.Windows.MessageBoxResult]::Yes) { return }
        }

        $active = Get-ActiveField $script:Fields
        $cfg = @{
            TemplatePath        = $ctl.txtTemplate.Text
            OutputPath          = $out
            ProjectName         = $ctl.txtProjectName.Text
            ProjectDescription  = $ctl.txtProjectDesc.Text
            DefaultEnvironment  = (Get-ComboText $ctl.cboEnv)
            DataFolderName      = $(if ([string]::IsNullOrWhiteSpace($ctl.txtDataFolder.Text)) { 'Data' } else { $ctl.txtDataFolder.Text })
            ConfigSheetName     = $(if ([string]::IsNullOrWhiteSpace($ctl.txtConfigSheet.Text)) { 'Settings' } else { $ctl.txtConfigSheet.Text })
            FileFilter          = $(if ([string]::IsNullOrWhiteSpace($ctl.txtFileFilter.Text)) { '*.xlsx' } else { $ctl.txtFileFilter.Text })
            Priority            = (Get-ComboText $ctl.cboPriority)
            QueueNameKey        = $ctl.txtQueueNameKey.Text
            QueueFolderKey      = $ctl.txtQueueFolderKey.Text
            Fields              = $active
            ReferenceExpression = (New-ReferenceExpression -Blocks @($script:RefSel) -TrimParts ([bool]$ctl.chkRefTrim.IsChecked))
            TruncateReference   = [bool]$ctl.chkRefTruncate.IsChecked
            ConfigItems         = @($script:ConfigRow)
        }

        # Die Generierung laeuft auf dem UI-Thread; das Fenster reagiert solange
        # nicht. IsEnabled verwirft in der Zwischenzeit angesammelte Klicks, damit
        # nicht direkt ein zweiter Lauf hinterherkommt. Kein IsEnabled auf
        # btnCreate selbst - das wuerde dessen Binding an chkConfirm ueberschreiben.
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        $window.IsEnabled = $false
        try   { $log = Invoke-DispatcherGeneration -Cfg $cfg }
        finally { $window.IsEnabled = $true }

        $stamp = Get-Date -Format 'dd.MM.yyyy HH:mm:ss'
        $lines = New-Object System.Collections.ArrayList
        [void]$lines.Add("Dispatcher erstellt am $stamp")
        [void]$lines.Add("Ziel: $out")
        [void]$lines.Add('')
        $nr = 0
        foreach ($entry in $log) { $nr++; [void]$lines.Add(("{0,3}. {1}" -f $nr, $entry)) }
        [void]$lines.Add('')
        [void]$lines.Add("Fertig - $nr Arbeitsschritte ohne Fehler abgeschlossen.")

        # Hinweise aus der Generierung nicht in der Liste untergehen lassen.
        $warn = @($log | Where-Object { $_ -like '*ACHTUNG*' })
        if ($warn.Count -gt 0) {
            [void]$lines.Add('')
            [void]$lines.Add("$($warn.Count) Hinweis(e) - bitte in Studio pruefen:")
            foreach ($w in $warn) { [void]$lines.Add("  - $w") }
        }

        $ctl.txtLog.Text = ($lines -join [Environment]::NewLine)
        Set-Status 'Dispatcher erfolgreich erzeugt.'
        if ($warn.Count -gt 0) {
            Show-Warn "Der Dispatcher wurde erstellt, es gibt aber $($warn.Count) Hinweis(e). Details stehen im Protokoll auf Maske 5."
        }
        $r = [System.Windows.MessageBox]::Show($window,
            "Der Dispatcher wurde erstellt:`r`n$out`r`n`r`nZielordner jetzt oeffnen?",
            'Fertig', 'YesNo', 'Information')
        if ($r -eq [System.Windows.MessageBoxResult]::Yes) { Start-Process explorer.exe $out }
    }
    catch {
        # Die Generierung ist nicht transaktional: bricht sie nach dem Kopieren ab,
        # bleibt ein vollstaendig aussehendes, aber halbfertiges Projekt liegen.
        # Ohne Markerdatei haette der Anwender es fuer brauchbar halten koennen.
        if ($out -and (Test-Path -LiteralPath $out)) {
            try {
                $marker = Join-Path $out 'GENERIERUNG_UNVOLLSTAENDIG.txt'
                $mtext  = "Die Generierung wurde am $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss') mit einem Fehler abgebrochen." + [Environment]::NewLine +
                          "Dieses Projekt ist unvollstaendig und darf nicht verwendet werden." + [Environment]::NewLine + [Environment]::NewLine +
                          $_.Exception.Message
                [System.IO.File]::WriteAllText($marker, $mtext, (New-Object System.Text.UTF8Encoding($false)))
            }
            catch { }
        }
        $ctl.txtLog.Text = "FEHLER" + [Environment]::NewLine +
                           "------------------------------------------------------------------" + [Environment]::NewLine +
                           $_.Exception.Message + [Environment]::NewLine + [Environment]::NewLine +
                           $_.ScriptStackTrace
        Show-Error "Die Erzeugung ist fehlgeschlagen:`r`n$($_.Exception.Message)"
    }
    finally { $window.Cursor = [System.Windows.Input.Cursors]::Arrow }
})

# ---------------------------------------------------------------------------
#  Initialisierung
# ---------------------------------------------------------------------------
Reset-MandatoryField
Reset-ConfigRow -Dev @{} -Prd @{}
Update-RefAvailable
Update-RefPreview

if ($TemplatePath) { $ctl.txtTemplate.Text = $TemplatePath; & $script:LoadTemplateConfig }
if ($OutputPath)   { $ctl.txtOutput.Text   = $OutputPath }
if ([string]::IsNullOrWhiteSpace($ctl.txtProjectName.Text)) { $ctl.txtProjectName.Text = 'Neuer Dispatcher' }
$ctl.txtProjectDesc.Text = 'Dispatcher - liest Quell-Excel und befuellt die Orchestrator-Queue.'

Update-Navigation
[void]$window.ShowDialog()
