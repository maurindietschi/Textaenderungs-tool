$folderPath = "Q:\AppTesting\QFTestFrameWork\QFTestDriver\Syrius\CEN_UNI"
$selection = $null

function showmenu {
  Clear-Host
  Write-Host "=== Vergleichs Tool ==="
  Write-Host "1. Option 1"
  Write-Host "2. Option 2"
  Write-Host "3. Durchsuchen und Datei auswählen"
  Write-Host "4. Beenden"
}

function Get-FilePath {
  Param (
    [string]$folderPath
  )

  $previousPaths = @()
  $currentPath = $folderPath

  while ($true) {
    Clear-Host
    $items = Get-ChildItem -Path $currentPath

    $index = 1
    $selection = 0

    Write-Host "Current path: $currentPath"
    foreach ($item in $items) {
      $itemName = $item.Name
      
      Write-Host "[$index] Name: $itemName"
      Write-Host "---------------"
      $index++
    }

    $selection = Read-Host "Wählen Sie eine Option (1 - $($index - 1))  oder 'z' um zum vorherigen Pfad zurückzukehren, oder geben Sie 'q' ein um abzubrechen."

    if ($selection -eq "q") {
      return $currentPath
    }
    elseif ($selection -eq "z") {
      if ($previousPaths.Count -gt 0) {
        $currentPath = $previousPaths[$previousPaths.Count - 1]
        $previousPaths = $previousPaths[0..($previousPaths.Count - 2)]
      }
    }
    elseif ([int]$selection -ge 1 -and [int]$selection -lt $index) {
      $selectedItem = $items[[int]$selection - 1]
      $selectedPath = $selectedItem.FullName

      if ($selectedItem.PSIsContainer) {
        $previousPaths += $currentPath
        $currentPath = $selectedPath
      }
      else {
        return $selectedPath
      }
    }
    else {
      Write-Host "Ungültige Auswahl!"
    }
  }
}

while ($selection -ne "4") {
  showmenu
  $selection = Read-Host "Wählen Sie eine Option aus (1-4)"

  switch ($selection) {
    "1" {
      Write-Host "Option 1 ausgewählt"
      Read-Host "Drücken Sie eine beliebige Taste, um fortzufahren..."
    }
    "2" {
      Write-Host "Option 2 ausgewählt"
      Read-Host "Drücken Sie eine beliebige Taste, um fortzufahren..."
    }
    "3" {
      Write-Host "Durchsuchen und Datei auswählen"
      $selectedFilePath = Get-FilePath -folderPath $folderPath

      Clear-Host
      Write-Host "Ausgewählter Dateipfad: $selectedFilePath"
      Read-Host "Drücken Sie eine beliebige Taste, um fortzufahren..."
    }
    "4" {

    }
    default {
      Write-Host "Ungültige Auswahl! Bitte geben Sie eine gültige Option ein."
      Read-Host "Drücken Sie eine beliebige Taste, um fortzufahren..."
    }
  }
}
