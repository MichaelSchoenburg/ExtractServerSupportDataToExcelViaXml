<#
.SYNOPSIS
    Aggregiert Inhalte aus der Spalte "MAC-Adresse" mehrerer XML-Tabellen aus mehreren ZIP-Dateien und speichert sie in einer CSV-Datei.

.DESCRIPTION
    Dieses Skript durchsucht ein angegebenes Verzeichnis nach XML-Dateien, extrahiert die Werte aus der angegebenen Spalte (standardmaessig "MAC-Adresse") aller enthaltenen Tabellen und speichert die aggregierten Ergebnisse in einer CSV-Datei.

.PARAMETER XMLFilesDirectory
    Das Verzeichnis, in dem sich die XML-Dateien befinden.
    Standardwert: "$PSScriptRoot\XML-Dateien"

.PARAMETER columnNameToSelect
    Der Name der Tabellenspalte, deren Werte extrahiert werden sollen.
    Standardwert: "MAC-Adresse"

.PARAMETER outputCsvPath
    Der Pfad zur Ausgabedatei (CSV), in der die aggregierten MAC-Adressen gespeichert werden.
    Standardwert: "$PSScriptRoot\aggregierte_mac_adressen.csv"

.EXAMPLE
    .\Skript.ps1
    Fuehrt das Skript mit den Standardwerten aus und speichert die aggregierten MAC-Adressen in einer CSV-Datei.

    .\Skript.ps1 -ZipOrdner "Pfad\zu\deinen\ZIPs" -CsvDatei "output.csv" -columnNameToSelect "MAC-Adresse"
    Fuehrt das Skript aus und gibt den Pfad zu den ZIP-Dateien und die gewuenschte CSV-Ausgabedatei an.
.LINK
    https://github.com/MichaelSchoenburg/ExtractServerSupportDataToExcelViaXml

.NOTES
    Autor: Michael Schoenburg
    Erstellt: 12.06.2025
#>

#region Parameter

param(
    [CmdletBinding()]
    # Pfad zum Export der Support-Daten
    [Parameter(
        Mandatory = $true,
        HelpMessage = "Dies ist der Pfad zum Ordner in welchem sich alle ZIP-Dateien, in welchen sich die XML-Datei 'viewer.XML' befindet befinden. In der Regel heißt der Ordner etwas mit *TSRLogs und muss aus der *TSRLogs.zip"
    )]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({
        if (-not ($_ -is [string] -and $_.Trim().Length -gt 0)) {
            throw "Der Pfad darf nicht leer sein."
        }
        if (-not (Test-Path $_ -PathType Container)) {
            throw "Der angegebene Pfad '$_' existiert nicht oder ist kein Ordner/Verzeichnis. Bitte nicht den direkten Pfad zur ZIP-Datei angeben, sondern den Ordner darueber. Dies ist der Pfad zum Ordner in welchem sich alle ZIP-Dateien, in welchen sich die XML-Datei 'viewer.XML' befindet befinden. In der Regel heißt der Ordner etwas mit *TSRLogs und muss aus der *TSRLogs.zip."
        }
        return $true
    })]
    [string]
    $ZipFilesDirectory = $PSScriptRoot,

    # Pfad zur Excel-Datei
    [Parameter(
        Mandatory = $false,
        HelpMessage = "Dies ist der Pfad zur Excel-Datei des Kunden, in welche die Daten eingetragen werden sollen. Wenn nicht angegeben, wird die erste gefundene Excel-Datei im aktuellen Verzeichnis verwendet."
    )]
    [ValidateNotNullOrEmpty()]
    # ValidateScript -> Falls ueberhaupt ein Wert angegeben wurde, pruefe, ob es ein gueltiger Pfad ist
    [ValidateScript({
        if ($null -ne $_) {
            if (-not (Test-Path $_)) {
                throw "Die Datei $_ konnte nicht gefunden werden. Bitte ueberpruefen Sie den Pfad."
            }
        }
        return $true
    })]
    [string]
    $PathToExcelFile,

    # Silent-Mode
    [Parameter(
        HelpMessage = "Im Silent-Mode wird Excel nicht angezeigt."
    )]
    [switch]
    $Silent = $false
)

#region Initialisierung

Add-Type -AssemblyName System.IO.Compression.FileSystem

#endregion
#region Funktionen



#endregion

# oeffne frueh eine Try-Schleife, um zu ermoeglich, dass falls zu irgendeinem Punkt 
# in der Ausfuehrung etwas schieflaufen sollte, das Aufraumen immer durchgefuehrt wird
try {
    #region Variablen

    # Ordner, in welchem die Extrahierten ZIP-Dateien und viewer.XML-Dateien gespeichert werden
    $baseTempDir = Join-Path -Path $env:TEMP -ChildPath "_ParseXMLFileContentsSkript"

    # Starte den Timer
    $timer = [System.Diagnostics.Stopwatch]::StartNew()

    # Pfad zur Excel-Datei durch Benutzer auswaehlen lassen, falls nicht als Parameter angegeben wurde
    if ($null -eq $PathToExcelFile -or $PathToExcelFile.Length -eq 0) {
        $ExcelFiles = (Get-ChildItem -Path "$PSScriptRoot" -Recurse -Filter "*.xlsx" -File).FullName
        if ($ExcelFiles.Count -eq 0) {
            throw "Es wurde keine Excel-Datei im Verzeichnis, wo das Skript gestartet wurde gefunden und auch keine Excel-Datei per Parameter angegeben. Bitte spezifizieren Sie welche Excel-Datei genutzt werden soll, indem Sie diese im Parameter -PathToExcelFile angeben oder die Excel-Datei in den Ordner verschieben, wo das Skript gestartet wird oder einen der Unterordner."
        } elseif ($ExcelFiles.Count -eq 1) {
            $title = $null
            $question = "Es wurde folgende Excel-Datei gefunden: '$($ExcelFiles)'. Moechten Sie die aus den Server-Support-Dateien extrahierten Daten in diese Excel-Datei schreiben? Waehlen Sie 'j' fuer ja oder 'a' fuer abbrechen. Bestaetigen Sie mit Enter."
            $a = [System.Management.Automation.Host.ChoiceDescription]::new( "&abbrechen", '' )
            $j = [System.Management.Automation.Host.ChoiceDescription]::new( "&ja", '' )
            $options = [System.Management.Automation.Host.ChoiceDescription[]]( $a, $j )
            $result = $host.ui.PromptForChoice( $title, $question, $options, 0 )

            if ($result -eq 1) {
                $PathToExcelFile = $ExcelFiles
                Write-Host "Sie haben 'Ja' gewaehlt. Das Skript wird die extrahierten Daten in die Excel-Datei '$($PathToExcelFile)' schreiben."
            } else {
                throw "Sie haben das Skript abgebrochen. Es wrude keine Excel-Datei per Parameter angegeben und die gefundene Excel-Datei wollten Sie nicht verwenden. Bitte spezifizieren Sie welche Excel-Datei genutzt werden soll, indem Sie diese im Parameter -PathToExcelFile angeben oder die Excel-Datei in den Ordner verschieben, wo das Skript gestartet wird oder einen der Unterordner."
            }
        } elseif ($ExcelFiles.Count -gt 1) {
            Write-Host "Es wurde mehrere Excel-Datei gefunden. In welche der Dateien moechten Sie die aus den Server-Support-Dateien extrahierten Daten schreiben?"

            $title = $null
            $question = "Waehlen Sie die Nummer der Excel-Datei aus oder waehlen Sie 'a' fuer 'abbrechen', wenn sie keine der Dateien verwenden wollen. Bestaetigen Sie mit Enter."
            
            $options = [System.Collections.Generic.List[System.Management.Automation.Host.ChoiceDescription]]::new()
            $options += [System.Management.Automation.Host.ChoiceDescription]::new( "&abbrechen", '' )
            
            for ($i = 0; $i -lt $ExcelFiles.Count; $i++) {
                Write-Host "[$($i)] - $($ExcelFiles[$i])"
                $options += [System.Management.Automation.Host.ChoiceDescription]::new( "&$i", "$([System.IO.Path]::GetFileName($ExcelFiles[$i]))" )
            }
            
            $result = $host.ui.PromptForChoice( $title, $question, $options, 0 )

            if ($result -ne 0) {
                $PathToExcelFile = $ExcelFiles[$result-1]
                Write-Host "Sie haben die Excel-Datei Nr. $($result) gewaehlt. Das Skript wird die extrahierten Daten in die Excel-Datei '$($PathToExcelFile)' schreiben."
            } elseif ($result -eq 0) {
                throw "Sie haben das Skript abgebrochen. Es wrude keine Excel-Datei per Parameter angegeben und die gefundene Excel-Datei wollten Sie nicht verwenden. Bitte spezifizieren Sie welche Excel-Datei genutzt werden soll, indem Sie diese im Parameter -PathToExcelFile angeben oder die Excel-Datei in den Ordner verschieben, wo das Skript gestartet wird oder einen der Unterordner."
            }
        } else {
            throw 'Unerwarteter Fehler.'
        }
    }

    #endregion

    #region ZIP-Dateien-Extraktion
    Write-Verbose "Schritt 1 von 6: Pruefe, ob ein temporaeres Verzeichnis fuer die Extraktion der ZIP- und XML-Dateien existiert..."
    if (-not (Test-Path $baseTempDir)) {
        Write-Verbose "Schritt 1 von 6: Dies wurde nicht gefunden. Erstelle ein temporaeres Verzeichnis fuer die Extraktion der ZIP- und XML-Dateien..."
        New-Item -ItemType Directory -Path $baseTempDir | Out-Null
    } else {
        Write-Verbose "Schritt 1 von 6: Wurde gefunden. Existiert bereits."
    }

    Write-Verbose "Schritt 2 von 6: Starte die Extraktion aller ZIP-Dateien im Verzeichnis '$($ZipFilesDirectory)'..."
    $zipFiles = Get-ChildItem -Path $ZipFilesDirectory -Filter "*.zip" -File -Recurse
    $XmlFiles = New-Object System.Collections.Generic.List[object]
    $n = 0

    foreach ($zip in $zipFiles) {
        $n++    
        $zipBaseName = [System.IO.Path]::GetFileNameWithoutExtension($zip.FullName)
        $tempDir = Join-Path -Path $baseTempDir -ChildPath $zipBaseName
        if (!(Test-Path -Path $tempDir)) {
            New-Item -ItemType Directory -Path $tempDir | Out-Null
        }

        # Write-Progress zum Anzeigen des Fortschritts
        $progressParams = @{
            Activity = "Schritt 2 von 6: Extrahiere ZIP-Dateien"
            Status   = "Extrahiere '$($zip.Name)'..."
            PercentComplete = ($n / $zipFiles.Count) * 100
        }
        Write-Progress @progressParams

        try {
            Write-Debug "Extrahiere ZIP-Datei: $($zip.FullName) nach $($tempDir)"
            [System.IO.Compression.ZipFile]::ExtractToDirectory($zip.FullName, $tempDir)

            $SubZipFiles = Get-ChildItem -Path $tempDir -Filter "*.zip" -File

            foreach ($subZip in $SubZipFiles) {
                $zipArchive = [System.IO.Compression.ZipFile]::OpenRead($subZip.FullName)
                foreach ($entry in $zipArchive.Entries) {
                    if ($entry.Name -eq 'sysinfo_DCIM_View.xml') {
                        $destPath = Join-Path $tempDir $entry.Name
                        Write-Debug "Extrahiere aus ZIP-Datei: $($subZip.FullName) XML-Datei: $($entry.Name) nach $($destPath)"
                        $entryStream = $entry.Open()
                        $fileStream  = [System.IO.File]::OpenWrite($destPath)
                        $entryStream.CopyTo($fileStream)
                        $fileStream.Close()
                        $entryStream.Close()
                        $XmlFiles.Add($destPath)
                    }
                }
                $zipArchive.Dispose()
            }
            
            Write-Debug "Alle ZIP-Dateien erfolgreich extrahiert."
        } catch {
            # Schmeisse nicht-terminierenden Fehler
            Write-Warning "Schritt 2 von 6: Fehler beim Extrahieren von '$($zip.FullName)': $($_.Exception.Message)"
        }
    }

    if (-not $XmlFiles -or $XmlFiles.Count -eq 0) {
        throw "Es wurde keine einzige XML-Dateie in den ZIP-Archiven gefunden. Somit koennen keine Daten gelesen werden. Skript wird beendet."
    } else {
        Write-Verbose "Schritt 2 von 6: Extraktion abgeschlossen. Insgesamt $($XmlFiles.Count) XML-Dateien gefunden und extrahiert."
    }

    #endregion

    #region XML-Parsing

    Write-Verbose "Schritt 3 von 6: Parse Daten aus den XML-Dateien aus..."

    $XmlsData = New-Object System.Collections.Generic.List[PSCustomObject]

    foreach ($xmlFile in $xmlFiles) {
        # Write-Progress zum Anzeigen des Fortschritts
        $progressParams = @{
            Activity = "Schritt 3 von 6: Parse XML-Dateien"
            Status   = "Verarbeite '$($xmlFile)'..."
            PercentComplete = ($XmlFiles.IndexOf($xmlFile) / $XmlFiles.Count) * 100
        }
        Write-Progress @progressParams

        # XML-Datei einlesen
        [xml]$xmlContent = Get-Content -Path $xmlFile

        # Alles aus System View
        $DCIM_SystemView = $xmlContent.SelectNodes('//INSTANCE[@CLASSNAME="DCIM_SystemView"]')
        $model = $DCIM_SystemView.Property.Where({ $_.Name -eq 'Model' }).Value
        $chassisServiceTag = $DCIM_SystemView.Property.Where({ $_.Name -eq 'chassisServiceTag' }).Value

        # Alles aus iDRACC Card View
        $DCIM_iDRACCardView = $xmlContent.SelectNodes('//INSTANCE[@CLASSNAME="DCIM_iDRACCardView"]')
        $ServernameOriginal = $DCIM_iDRACCardView.Property.Where({ $_.Name -eq 'DNSRacName' }).Value
        $Servername = $ServernameOriginal.Replace('-idrac', '')
        $ServerMacAddress = $DCIM_iDRACCardView.Property.Where({ $_.Name -eq 'PermanentMACAddress' }).Value

        # Alle Ethernet NICs
        $DCIM_NICView = $xmlContent.SelectNodes('//INSTANCE[@CLASSNAME="DCIM_NICView"]')
        $EthernetNics = foreach ($nic in $DCIM_NICView) {
            $PoductName = $nic.Property.Where({ $_.Name -eq 'ProductName' }).Value.Split('-')[0]
            [PSCustomObject]@{
                DeviceDescription   = $nic.Property.Where({ $_.Name -eq 'DeviceDescription' }).Value
                MACAddress          = $nic.Property.Where({ $_.Name -eq 'CurrentMACAddress' }).Value
                ProductName         = $PoductName
                PortSpeed           = if ($PoductName -like '*Gigabit*') {
                                        '1 G'
                                    } elseif ($PoductName -match ' \dx(100G) ') {
                                        '100 G QSFP'
                                    } elseif ($PoductName -match ' \dx(10G|25G) ') {
                                        '10/25 GbE SFP'
                                    } else {
                                        $null
                                    }
            }
        }

        # Alle InfiniBand NICs
        if ($xmlContent.SelectNodes('//INSTANCE[@CLASSNAME="DCIM_InfiniBandView"]')) {
            $DCIM_InfiniBandView = $xmlContent.SelectNodes('//INSTANCE[@CLASSNAME="DCIM_InfiniBandView"]')
            $InfiniBandNics = foreach ($nic in $DCIM_InfiniBandView) {
                [PSCustomObject]@{
                    DeviceDescription   = $nic.Property.Where({ $_.Name -eq 'DeviceDescription' }).Value
                    MACAddress          = $nic.Property.Where({ $_.Name -eq 'CurrentMACAddress' }).Value
                    ProductName         = $nic.Property.Where({ $_.Name -eq 'ProductName' }).Value.Split('-')[0]
                    PortSpeed           = 'InfiniBand'
                }
            }
        } else {
            $InfiniBandNics = $null
        }

        # Ethernet- und Infiniband-NICs in eine Tabelle kombinieren
        $AllNics = $EthernetNics + $InfiniBandNics

        # Alle physical Disks
        if ($xmlContent.SelectNodes('//INSTANCE[@CLASSNAME="DCIM_PhysicalDiskView"]')) {
            $DCIM_PhysicalDiskView = $xmlContent.SelectNodes('//INSTANCE[@CLASSNAME="DCIM_PhysicalDiskView"]')
            $PhysicalDisks = foreach ($disk in $DCIM_PhysicalDiskView) {
                [PSCustomObject]@{
                    Device              = $disk.Property.Where({ $_.Name -eq 'Model' }).Value
                    Slot                = $disk.Property.Where({ $_.Name -eq 'DeviceDescription' }).Value
                    DiskType            = $disk.Property.Where({ $_.Name -eq 'Model' }).Value
                    Serialnumber        = $disk.Property.Where({ $_.Name -eq 'SerialNumber' }).Value
                    SasAddress          = $disk.Property.Where({ $_.Name -eq 'SASAddress' }).Value
                }
            }
        } else {
            $PhysicalDisks = $null
        }

        # Erstelle ein PSCustomObject, um alle Daten zu speichern
        $xmlData = [PSCustomObject]@{
            Model               = $model
            ChassisServiceTag   = $chassisServiceTag
            Servername          = $Servername
            ServerMacAddress    = $ServerMacAddress
            NICs                = $AllNics
            PhysicalDisks       = $PhysicalDisks
        }

        # Alle Daten dieser XML-Datei in die uebergeordnete Tabelle hinzufuegen
        $XmlsData += $xmlData
    }

    Write-Verbose "Schritt 3 von 6: Alle XML-Dateien geparst."

    #endregion

    
    #region Excel-Initialisierung

    Write-Verbose "Schritt 4 von 6: Initialisiere Excel..."
    $excel = New-Object -ComObject Excel.Application

    if ($Silent) {
        Write-Verbose "Schritt 4 von 6: Das Skript wurde nicht im Silent-Modus gestartet. Excel wird angezeigt. Wenn du das nicht moechtest, starte das Skript das naechste Mal mit dem Parameter -Silent."
        $excel.Visible = $false
    } else {
        Write-Verbose "Schritt 4 von 6: Das Skript wurde im Silent-Modus gestartet. Excel wird nicht angezeigt. Wenn du die Excel-Datei live sehen moechtest, waehrend die Daten eingetragen werden, starte das Skript das naechste Mal ohne den Parameter -Silent."
        $excel.Visible = $true
    }
    
    # Hole den absoluten Pfad und Dateinamen aus $PathToExcelFile
    $excelFullPath = Resolve-Path -Path $PathToExcelFile | Select-Object -ExpandProperty Path
    $workbook = $excel.Workbooks.Open($excelFullPath)

    Write-Verbose "Schritt 4 von 6: Excel-Initialisierung abgeschlossen."

    #endregion

    #region Excel-Schreiben

    Write-Verbose "Schritt 5 von 6: Schreibe Daten in Excel-Datei..."

    foreach ($myXmlData in $XmlsData) {
        $progressParams = @{
            Activity = "Schritt 5 von 6: Schreibe Daten in Excel-Datei..."
            Status   = "Schreibe Daten von '$($myXmlData.Servername)' in Excel..."
            PercentComplete = ($XmlsData.IndexOf($myXmlData) / $XmlsData.Count) * 100
        }
        Write-Progress @progressParams

        Write-Verbose "--------------------------------------------------------------------------"
        Write-Verbose "Arbeitsblatt 10 auswaehlen..."
        $worksheet = $workbook.Worksheets.Item(10)
        $lastRow = $worksheet.UsedRange.Rows.Count

        Write-Verbose "Trage Daten in Excel (Arbeitsblatt 'Geraete Interface MAC') ein..."
        $ServerRow = $null
        for ($i = 3; $i -le $lastRow; $i++) {
            $cellValue = $worksheet.Cells.Item($i, 3).Value2
            if ($cellValue -eq $myXmlData.Servername) {
                Write-Verbose "Servername '$($myXmlData.Servername)' gefunden in Zeile $i"
                $ServerRow = $i
                break
            }
        }
        if (-not $ServerRow) {
            # Suche die naechste komplett leere Zeile ab Zeile 3
            Write-Verbose "Finde freie Zeile und trage Servername '$($myXmlData.Servername)' ein..."
            for ($i = 3; $i -le ($lastRow + 1000); $i++) {
                $rowIsEmpty = $true
                for ($col = 1; $col -le $worksheet.UsedRange.Columns.Count; $col++) {
                    if ($worksheet.Cells.Item($i, $col).Value2) {
                        $rowIsEmpty = $false
                        break
                    }
                }
                if ($rowIsEmpty) {
                    $ServerRow = $i
                    Write-Verbose "Trage Servername '$($myXmlData.Servername)' in neue Zeile $ServerRow Spalte 3 ein..."
                    $worksheet.Cells.Item($ServerRow, 3).Value2 = $myXmlData.Servername
                    Write-Verbose "Servername '$($myXmlData.Servername)' in neue Zeile $ServerRow eingetragen."
                    break
                }
            }
            if (-not $ServerRow) {
                throw "Keine freie Zeile gefunden, um den Servernamen '$($myXmlData.Servername)' einzutragen."
            }
        }

        # Geraet
        Write-Verbose "Trage Model '$($myXmlData.Model)' in Zeile $ServerRow, Spalte 4 mit der ueberschrift 'Geraet' ein..."
        $worksheet.Cells.Item($ServerRow, 4).Value2 = $myXmlData.Model

        # MAC Address > MAC
        Write-Verbose "Trage iDrac-MAC-Adresse '$($myXmlData.ServerMacAddress)' in Zeile $ServerRow, Spalte 6 ein..."
        $worksheet.Cells.Item($ServerRow, 6).Value2 = $myXmlData.ServerMacAddress

        # MAC Address > User / PW
        Write-Verbose "Trage Benutzer und Passwort in Zeile $ServerRow, Spalte 7 ein..."
        $worksheet.Cells.Item($ServerRow, 7).Value2 = "root/calvin"

        # Firmware > Port 1
        Write-Verbose "Finde die MAC-Adresse von Embedded 1 Port 1..."
        $FirmwarePort1MacAddress = $myXmlData.Nics.Where({ $_.DeviceDescription -like "*Embedded*NIC 1*Port 1*" })."MACAddress"
        Write-Verbose "Trage MAC-Adresse '$FirmwarePort1MacAddress' von Embedded NIC 1 Port 1 in Zeile $ServerRow, Spalte 8 ein..."
        $worksheet.Cells.Item($ServerRow, 8).Value2 = $FirmwarePort1MacAddress

        # Firmware > Port 2
        Write-Verbose "Finde die MAC-Adresse von Embedded NIC 1 Port 2..."
        $FirmwarePort2MacAddress = $myXmlData.Nics.Where({ $_.DeviceDescription -like "*Embedded*NIC 1*Port 2*" })."MACAddress"
        Write-Verbose "Trage MAC-Adresse '$FirmwarePort2MacAddress' von Embedded NIC 1 Port 2 in Zeile $ServerRow, Spalte 9 ein..."
        $worksheet.Cells.Item($ServerRow, 9).Value2 = $FirmwarePort2MacAddress

        # Netzwerkkarten 100 G QSFP
        Write-Verbose "Bereite 100 G QSFP-Netzwerkkarten vor..."
        $Nics100G = $myXmlData.Nics.Where({ $_.PortSpeed -eq "100 G QSFP" })
        $j = 0
        foreach ($nic in $Nics100G) {
            Write-Verbose "Trage Name von $($nic.DeviceDescription) in Zeile $($ServerRow), Spalte $($j) ein..."
            $worksheet.Cells.Item($ServerRow, 10 + $j).Value2 = $nic.DeviceDescription
            $j += 1
            Write-Verbose "Trage MAC-Adresse von $($nic.DeviceDescription) in Zeile $($ServerRow), Spalte $($j) ein..."
            $worksheet.Cells.Item($ServerRow, 10 + $j).Value2 = $nic.MACAddress
            $j += 1
        }

        # Netzwerkkarten 10/25 GbE SFP
        $Nics1025G = $myXmlData.Nics.Where({ $_.PortSpeed -eq "10/25 GbE SFP" })
        $n = 0
        foreach ($nic in $Nics1025G) {
            Write-Verbose "Trage Name von $($nic.DeviceDescription) in Zeile $($ServerRow), Spalte $($n) ein..."
            $worksheet.Cells.Item($ServerRow, 26 + $n).Value2 = $nic.DeviceDescription
            $n += 1
            Write-Verbose "Trage MAC-Adresse von $($nic.DeviceDescription) in Zeile $($ServerRow), Spalte $($n) ein..."
            $worksheet.Cells.Item($ServerRow, 26 + $n).Value2 = $nic.MACAddress
            $n += 1
        }

        # Netzwerkkarten InfiniBand
        Write-Verbose "Bereite InfiniBand-Netzwerkkarten vor..."
        $NicsInfiniBand = $myXmlData.Nics.Where({ $_.PortSpeed -eq "InfiniBand" })
        $m = 0
        foreach ($nic in $NicsInfiniBand) {
            Write-Verbose "Trage Name von $($nic.DeviceDescription) in Zeile $($ServerRow), Spalte $($m) ein..."
            $worksheet.Cells.Item($ServerRow, 42 + $m).Value2 = $nic.DeviceDescription
            $m += 1
            Write-Verbose "Trage MAC-Adresse von $($nic.DeviceDescription) in Zeile $($ServerRow), Spalte $($m) ein..."
            $worksheet.Cells.Item($ServerRow, 42 + $m).Value2 = $nic.MACAddress
            $m += 1
        }

        # Physical Disks / HW&Disk Serial Nr.
        Write-Verbose "----------------------------------------"
        Write-Verbose "Trage Daten in Excel (Arbeitsblatt 'HW&Disk Serial Nr.') ein..."
        $worksheet = $workbook.Worksheets.Item(9)
        $lastRow = $worksheet.UsedRange.Rows.Count
        # Setze Hintergrundfarbe der Zellen (Spalten 3 bis 7 in der neuen Zeile) auf Grau
        for ($col = 1; $col -le 10; $col++) {
            $cell = $worksheet.Cells.Item($lastRow + 1, $col)
            $cell.Interior.ColorIndex = 15  # 15 = Grau (Excel Standard)
        }
        # Server-Eckdaten
        $worksheet.Cells.Item($lastRow + 2,3).Value2 = $myXmlData.Servername
        $worksheet.Cells.Item($lastRow + 2,3).Font.Bold = $true
        $worksheet.Cells.Item($lastRow + 2,7).Value2 = $myXmlData.ChassisServiceTag
        $worksheet.Cells.Item($lastRow + 2,7).Font.Bold = $true

        if ($null -ne $myXmlData.PhysicalDisks -and $myXmlData.PhysicalDisks.Count -gt 0) {
            # Disks
            $i = 0
            foreach ($pd in $myXmlData.PhysicalDisks) {
                $i++

                # Spalte: Geraet
                $worksheet.Cells.Item($lastRow + 2 + $i,4).Value2 = $pd.Device
                
                # Spalte: Disk-Fach/Slot
                # Fach parsen
                if ($pd.Slot -like "*Backplane*") {
                    $DiskType = "Slot"
                } 
                if ($pd.Slot -like "*BOSS*") {
                    $DiskType = "BOSS"
                }

                # Slot parsen
                $Slot = [regex]::Matches($pd.Slot, "Disk \d+").Value
                $Slot = [regex]::Matches($Slot, "\d+").Value
                $Slot = [int]$Slot

                # Fuehrende Null hinzufuegen
                if ([int]$Slot -lt 10) {
                    $Slot = "{0:D2}" -f [int]$Slot
                }

                $worksheet.Cells.Item($lastRow + 2 + $i,5).Value2 = "$($DiskType) $($Slot)"

                # Spalte: Disk Type
                $worksheet.Cells.Item($lastRow + 2 + $i,6).Value2 = $pd.Device

                # Spalte: Seriennummer
                $worksheet.Cells.Item($lastRow + 2 + $i,7).Value2 = $pd.Serialnumber

                # Spalte: SAS Address
                $worksheet.Cells.Item($lastRow + 2 + $i,8).Value2 = $pd.SasAddress
            }
        }
    }

    Write-Verbose "Schritt 5 von 6: Alle Daten in die Excel-Datei uebertragen..."

    #endregion

} finally {
    #region Aufraeumen

    Write-Verbose "Schritt 6 von 6: Schliesse Excel-Datei und raeume extrahierte Dateien auf..."

    # Excel schliessen
    if ($workbook) {
        # Speichern der Excel-Datei
        $workbook.Save()
        
        $workbook.Close()
    }
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }

    # Loesche die temporaer angelegten Ordner inkl. Dateien darin
    try {
        if (Test-Path -Path $baseTempDir) {
            Remove-Item -Path $baseTempDir -Recurse -ErrorAction Stop
        }
    } catch {
        # Schmeisse nicht-terminierenden Fehler
        Write-Warning "Konnte temporaeren Ordner '$baseTempDir' nicht loeschen: $($_.Exception.Message)"
    }

    # Ausgeben, wie lange das Skript gebraucht hat
    $timer.Stop()
    Write-Host "Das Skript wurde in $($timer.Elapsed.ToString("hh\:mm\:ss")) Stunden abgeschlossen."
}

#endregion