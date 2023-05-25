# Pfad zu den Excel-Dateien
$quelleDatei = "C:\tmp\Textaenderungen_an_Adcubum_Syrius-GUI.xlsx"
$zielDatei = "Q:\AppTesting\QFTestFrameWork\QFTestDriver\Syrius\CEN_UNI\LEFK\CEN_UNI_LEFK_AQU_Test.xlsx"

# Erstelle Excel-Objekte
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

# Öffne die Quelldatei
$workbookQuelle = $excel.Workbooks.Open($quelleDatei)
$worksheetQuelle = $workbookQuelle.ActiveSheet
$lastRowQuelle = $worksheetQuelle.UsedRange.Rows.Count
$columnDRange = $worksheetQuelle.Range("D16:D$lastRowQuelle")

# Öffne die Zieldatei
$workbookZiel = $excel.Workbooks.Open($zielDatei)

# Durchlaufe alle Blätter in der Zieldatei ab dem zweiten Blatt
for ($i = 2; $i -le $workbookZiel.Worksheets.Count; $i++) {
    $worksheetZiel = $workbookZiel.Worksheets.Item($i)
    $rangeZiel = $worksheetZiel.UsedRange
    $lastRowZiel = $rangeZiel.Rows.Count

    # Durchlaufe die Werte in der Spalte D der Quelldatei
    foreach ($cell in $columnDRange.Cells) {
        $wert = $cell.Value2

        # Überspringe den Wert, wenn er leer ist
        if ([string]::IsNullOrEmpty($wert)) {
            continue
        }

        # Durchsuche jede Zelle in der Zieldatei nach dem Wert, ab der zweiten Zeile
        $treffer = $rangeZiel.Offset(1, 0).Resize($lastRowZiel-1).Find($wert)

        # Wenn der Wert gefunden wurde, gib eine Ausgabe aus
        if ($treffer) {
            $zielTabelle = $worksheetZiel.Name
            $zielZelle = $treffer.Address
            $quelleTabelle = $worksheetQuelle.Name
            $quelleZelle = $cell.Address
            Write-Host "Wert '$wert' wurde in der Tabelle '$zielTabelle', Zelle $zielZelle, gefunden. Ursprung: Tabelle '$quelleTabelle', Zelle $quelleZelle."
        }
    }
}

# Schließe die Excel-Objekte
$workbookQuelle.Close()
$workbookZiel.Close()
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheetQuelle) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheetZiel) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbookQuelle) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbookZiel) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()