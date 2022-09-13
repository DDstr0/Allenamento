<# Funzioni del programma: 
  - Selezione multipla CSV con GUI
  - Controllo validità input
  - Fusione degli export
  - Rimozione oggetti ripetuti con counter (opzionale)
  - Ordinamento degli export in base a proprietà scelta
  - Save file dialog

   Funzioni da implementare: 
  - 
   Note:
  - 

   Bug:
  - 
#>

[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.InitialDirectory = "$env:USERPROFILE"
$OpenFileDialog.Multiselect = $true
$OpenFileDialog.Title = "Seleziona gli export in CSV"
$OpenFileDialog.Filter = "Export (*.csv)|*.csv"

If ($OpenFileDialog.ShowDialog() -eq "Cancel"){
    Write-Warning "Nessun file selezionato."
    Exit
}


$OpenFileDialog.FileNames | ForEach-Object {
    [array]$temp = Get-Content -Path "$_" -Encoding "UTF8"

    If ($temp -eq $null){
        Write-Host "$_ : " -NoNewLine; Write-Host "EMPTY" -ForegroundColor "Red"
        Exit
    } ElseIf ($temp[0].StartsWith("#TYPE ") -eq $false){
        Write-Host "$_ : " -NoNewLine; Write-Host "UNSUPPORTED" -ForegroundColor "Red"
        Exit
    } Else {
        Write-Host "$_ : " -NoNewLine; Write-Host "OK" -ForegroundColor "Green"
    }
}; Write-Host ""



$data = @()
$OpenFileDialog.FileNames | ForEach-Object {
    $data += Import-CSV -Path "$_" -Encoding "UTF8"
}



If ((Read-Host -Prompt "Rimuovere duplicati? [Y/N]").ToLower() -eq "y"){

    $Duplicates = ($data | Group-Object -Property "NomeFile" | Where-Object -Property "Count" -GT 1).Count

    $temp = @()
    $data | Group-Object -Property "NomeFile" | ForEach-Object {$temp += $_.Group[0]}
    $data = $temp

    If ($Duplicates -eq 1){Write-Warning "1 ripetizione rimossa."} 
    Else {Write-Warning "$Duplicates ripetizioni rimosse."}
} Write-Host ""



$ErrorActionPreference = 'SilentlyContinue'
For ($i = 0; $i -le $data.Count - 1; $i++){
    Write-Progress -Activity "Parsing datetime and timespan..." -Status ( [string]( [math]::Round($i/($data.Count - 1)*100, 2) ) + "%" ) -PercentComplete ($i/($data.Count - 1)*100)
    $data[$i].TimeFromName = [datetime]::ParseExact($data[$i].TimeFromName, "dd/MM/yyyy HH:mm:ss", $null)
    $data[$i].Duration = [timespan]::Parse($data[$i].Duration)
} Write-Progress -Activity "Parsing datetime and timespan..." -Complete
$ErrorActionPreference = 'Continue'


$Properties = ($data | Get-Member -MemberType "NoteProperty").Name
Write-Host "Ordina in base a: " -ForegroundColor "Green"

For ($i = 0; $i -le $Properties.Count - 1; $i++){
    Write-Host ("[$i] " + $Properties[$i])
} Write-Host ""

$ErrorActionPreference = 'SilentlyContinue'
Do {[int]$Choice = Read-Host -Prompt "==>"
} While($Choice -lt 0 -or $Choice -gt ($Properties.Count - 1))

Write-Host ""

Write-Host "In ordine: " -ForegroundColor "Green"
Write-Host "[0] Crescente"; Write-Host "[1] Decrescente"; Write-Host ""

Do {[int]$temp = Read-Host -Prompt "==>"
} While($temp -lt 0 -or $temp -gt 1)
$ErrorActionPreference = 'Continue'

$Criteria = @{Property = $Properties[$Choice]; Descending = [bool]$temp}

$data = $data | Sort-Object @Criteria



$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$SaveFileDialog.InitialDirectory = "$env:USERPROFILE\Desktop"
$SaveFileDialog.Filter = "Merged export CSV (*.csv)| *.csv"
$OpenFileDialog.Title = "Salva l'export in CSV"

If ($SaveFileDialog.ShowDialog() -eq "OK"){
    $data | Export-CSV -Path $SaveFileDialog.FileName -Encoding "UTF8"
} Else {
    Write-Warning "Export non salvato."
}