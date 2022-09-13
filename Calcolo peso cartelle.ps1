#I valori nella proprietà "Length" sono espressi in Byte
#Bug: Se il programma viene caricato da file ps1 ci sono problemi con la codifica e la proprieta "Unità" di $data

[System.Reflection.Assembly]::LoadWithPartialName("System.windows.Forms") | Out-Null
$directory = New-Object System.Windows.Forms.FolderBrowserDialog
$directory.RootFolder = "Desktop"; $directory.ShowNewFolderButton = $false

If ($directory.ShowDialog() -eq "Cancel"){Write-Warning "Nessuna cartella selezionata."; Exit}

$Units = @("KB", "MB", "GB")

Write-Host "Seleziona l'unità di misura: "; Write-Host ""

For ($i = 0; $i -le $Units.Count - 1; $i++){
    Write-Host ("[$i] " + $Units[$i])
} Write-Host ""

Do{
    $Selected = Read-Host -Prompt "==>"
} While($Selected -lt 0 -or $Selected -gt $Units.Count - 1)

$Selected = $Units[$Selected]

$data = @()
Get-ChildItem -Path $directory.SelectedPath -Directory | ForEach-Object {
    $temp = 0

    Get-ChildItem -Path $_.FullName -File -Recurse | ForEach-Object {$temp += $_.Length/([string]1 + "$Selected"); <#Write-Host ("DEBUG: " + $_.Length + " Bytes")#>}
    $temp = [math]::Round($temp, 2)
    #Pause #Debug

    $data += New-Object -TypeName psobject -Property @{NomeCompleto = $_.FullName; Nome = $_.Name; Dimensione = $temp; Unità = $Selected}
}

$data = $data | Select-Object "NomeCompleto", "Nome", "Dimensione", "Unità"
$data