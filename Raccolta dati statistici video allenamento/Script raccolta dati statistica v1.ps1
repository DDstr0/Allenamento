<# Dati da estrarre:
- Data di creazione/modifica (se inesatta estrapola dal nome del file)
- Durata file (tieni conto delle foto)

Rappresentare con tabella e grafici su Excel:
- Quantità  di foto e video registrati per ogni giorno
- Quanto tempo mi sono allenato in un giorno
- Rappresentare l'aumento della durata degli allenamenti su un grafico 

https://social.technet.microsoft.com/Forums/lync/en-US/bad2dbb1-5deb-48b8-8f8c-45e2b353dba0/how-do-i-get-video-file-duration-in-powershell-script

# Bug: a volte il nome file e la data di modifica non coincidono; BUG: è probabile che si manifesti perché: (1. $data non viene dichiarato come array), (2. Il Select-Object prende solo una proprietà), (SICURO 3. '$data = $data.Name.Replace("-", "")')
  Possibile fix: crea fin da subito un altro array diverso da $data in cui conservi tutte le informazioni corrette in una NoteProperty

#Nota: 	Tieni conto che lo script va eseguito con tutti i file in possesso, aggiungi i dati ad un database
	Fa in modo che la proprietà  "Duration" non sia di tipo String ma Timespan - OK
	Ricorda di esportare in CSV
#>

[System.Reflection.Assembly]::LoadWithPartialName("System.windows.Forms") | Out-Null
$directory = New-Object System.Windows.Forms.FolderBrowserDialog
$directory.RootFolder = "Desktop"; $directory.ShowNewFolderButton = $false

$directory.ShowDialog() | Out-Null

If ($directory.SelectedPath -eq ""){Write-Warning "No directory selected."; $directory.SelectedPath = ""; Exit}



$data = @(Get-ChildItem -Path $directory.SelectedPath -Recurse -File -Include ("*.mp4", "*.jpg") | Sort-Object -Property "LastWriteTime" | Select-Object -Property Name, LastWriteTime)



$Length = @()
Get-ChildItem -Path $directory.SelectedPath -Recurse -File -Include ("*.mp4", "*.jpg") | Sort-Object -Property "LastWriteTime" | ForEach-Object {

    Write-Progress -Activity "Acquisendo durata video..."

    $Folder = $_.DirectoryName
    $File = $_.Name
    $LengthColumn = 27
    $objShell = New-Object -ComObject Shell.Application 
    $objFolder = $objShell.Namespace($Folder)
    $objFile = $objFolder.ParseName($File)
    $Length += $objFolder.GetDetailsOf($objFile, $LengthColumn)
}

$total = 0
For ($i = 0; $i -le $data.Count - 1; $i++){
    Write-Progress -Activity "Registrando dati nella proprietà 'Duration'..." -Status ( [string][math]::Round( ((100 / $data.Count) * $i), 2 ) + " completato" ) -PercentComplete ( (100 / $data.Count) * $i )
    If ($Length[$i] -ne ""){$data[$i] | Add-Member -MemberType NoteProperty -Name "Duration" -Value ([timespan]$Length[$i])} 
                      Else {$data[$i] | Add-Member -MemberType NoteProperty -Name "Duration" -Value ([timespan]::Parse("00:00:00"))}

    $total += $data.Duration[$i]
}

$total = $total | Select-Object -Property "TotalDays", "TotalHours", "TotalMinutes", "TotalSeconds" 

Write-Host "Totale tempo registrato:" -ForegroundColor Green
$total

$data | Out-GridView -Title 'Output array $data'