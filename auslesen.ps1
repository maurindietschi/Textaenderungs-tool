# Excel-Objekt erstellen
$excel = New-Object -ComObject Excel.Application

# Excel-Datei öffnen
$workbook = $excel.Workbooks.Open("Q:\AppTesting\QFTestFrameWork\QFTestDriver\Syrius\CEN_UNI\LEFK\CEN_UNI_LEFK_INN_Test.xlsx")

# Arbeitsblatt auswählen (z. B. das erste Arbeitsblatt)
$worksheet = $workbook.Worksheets.Item(1)

# Zellenwerte auslesen
$cellValue = $worksheet.Cells.Item(1, 1).Value()

# Wert anzeigen
Write-Host "Wert in Zelle A1: $cellValue"

# Excel-Datei schließen und Excel-Objekt aufräumen
$workbook.Close()
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null