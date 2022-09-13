<# Funzioni del programma: 
  - Selezione multipla CSV con GUI
  - Controllo validità input
  - Binding degli export
  - Rimozione oggetti ripetuti con counter (opzionale)
  - Ordinamento degli export in base a proprietà scelta
  - Genera nuovo export

   Funzioni da implementare: 

   Note:
  - Per la rimozione delle ripetizioni, potrei: 1. Conservare in una variabile temporanea l'oggetto ripetuto, 2. Rimuovere tutte le copie e l'originale, 3. Appendere alla fine l'originale e riordinare l'array con i criteri selezionati

   Bug:
  - A linea 62: dopo una modifica esce un errore di bad enumeration (previsto nella versione 1 dello script)
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
    Write-Progress -Activity "Binding exports..."
    $data += Import-CSV -Path "$_" -Encoding "UTF8"
}



If ((Read-Host -Prompt "[!] La seguente operazione può richiedere parecchio tempo. Rimuovere duplicati? [Y/N]").ToLower() -eq "y"){
    [System.Collections.ArrayList]$data = $data

    $Duplicates = 0
    $Done = 0

    $data | ForEach-Object {
        Write-Progress -Activity ("Cercando duplicati di '" + $_.NomeFile + "'") -Status ([string]([math]::Round(($Done/($data.Count - 1))*100, 2)) + "%") -PercentComplete (($Done/($data.Count - 1))*100)
        $Actual = $_
        $Count = 0

        For ($i = 0; $i -le $data.Count - 1; $i++){
            #Write-Progress -Activity ("Cercando duplicati di '" + $_.NomeFile + "'") -Status ([string]([math]::Round(($i/($data.Count - 1))*100, 2)) + "%") -PercentComplete (($i/($data.Count - 1))*100)
            If ($data[$i].NomeFile -eq $_.NomeFile){$Count++}
        }

        If ($Count -gt 1){
            $data.Remove($_) #Rimuove tutte le copie
            $data += $Actual #Appende all'ultimo

            $Duplicates++
        }

    Start-Sleep -Milliseconds 200
    $Done++
    }

    If ($Duplicates -eq 1){Write-Warning "1 ripetizione rimossa."} 
    Else {Write-Warning "$Duplicates ripetizioni rimosse."}
} Write-Host ""



$ErrorActionPreference = 'SilentlyContinue'
For ($i = 0; $i -le $data.Count - 1; $i++){
    Write-Progress -Activity "Parsing datetime and timespan..." -Status ( [string]( [math]::Round($i/($data.Count - 1)*100, 2) ) + "%" ) -PercentComplete ($i/($data.Count - 1)*100)
    $data[$i].TimeFromName = [datetime]::ParseExact($data[$i].TimeFromName, "dd/MM/yyyy HH:mm:ss", $null)
    $data[$i].Duration = [timespan]::Parse($data[$i].Duration)
} $ErrorActionPreference = 'Continue'


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



[System.Reflection.Assembly]::LoadWithPartialName("System.windows.Forms") | Out-Null
$directory = New-Object System.Windows.Forms.FolderBrowserDialog
$directory.RootFolder = "Desktop"; $directory.Description = "Seleziona cartella di output"

If ($directory.ShowDialog() -eq "Cancel"){
    $data | Export-CSV -Path "$env:USERPROFILE\Desktop\Export.csv" -Encoding "UTF8"
    Write-Warning "'Export.csv' generato in $env:USERPROFILE\Desktop"
} Else {
    $data | Export-CSV -Path ($directory.SelectedPath + "\Export.csv") -Encoding "UTF8"
}