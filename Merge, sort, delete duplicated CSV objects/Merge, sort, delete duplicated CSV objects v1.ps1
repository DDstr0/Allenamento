<# Funzioni del programma: 
  - Selezione multipla con GUI
  - Controllo validità input
  - Binding degli export
  - Ordinamento degli export in base a proprietà scelta

   Funzioni da implementare: 
  - Rimozione oggetti ripetuti con counter
  - Genera nuovo export

   Note:
  - Per la rimozione delle ripetizioni, potrei: 1. Conservare in una variabile temporanea l'oggetto ripetuto, 2. Rimuovere tutte le copie e l'originale, 3. Appendere alla fine l'originale e riordinare l'array con i criteri selezionati

   Bug:
  - A linea 102, dopo aver rimosso un elemento dall'array, nel caso in cui ci sia più di una ripetizione, è necessario decrementare il contatore tante volte quante le ripetizioni dopo la prima
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
    $temp = Get-Content -Path "$_" -Encoding "UTF8"

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


[System.Collections.ArrayList]$data = $data #Per avere la dimensione variabile e poter eventualmente cancellare le ripetizioni

$data | ForEach-Object {
    $IndexesToRemove = @()

    For ($i = 0; $i -le $data.Count - 1; $i++){
        If ($data[$i] -eq $_.NomeFile){$IndexesToRemove += $i}
    }

    If ($IndexesToRemove.Count -gt 1){
        #Spiegazione: parto da 1 e arrivo al totale di elementi da rimuovere - 1 (in modo da preservare il primo elemento dall'alto), e quando rimuovo da $data
        #uso come indice il valore negativo del contatore, in modo da eliminare dal basso l'oggetto dell'array
        For ($i = 1; $i -le $IndexesToRemove.Count - 1; $i++){
            $data.RemoveAt($IndexesToRemove[-$i])
        }
    }
}