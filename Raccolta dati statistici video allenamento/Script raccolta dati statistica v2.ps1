<# Dati da estrarre:
- Data di creazione/modifica (se inesatta estrapola dal nome del file) - OK
- Durata file (tieni conto delle foto) - OK

Rappresentare con tabella e grafici su Excel:
- Quantità di foto e video registrati per ciascun giorno
- Quanto tempo mi sono allenato in ciascun giorno
- Rappresentare l'aumento della durata degli allenamenti su un grafico 

https://social.technet.microsoft.com/Forums/lync/en-US/bad2dbb1-5deb-48b8-8f8c-45e2b353dba0/how-do-i-get-video-file-duration-in-powershell-script

#Nota: 	Tieni conto che lo script va eseguito con tutti i file in possesso, aggiungi i dati ad un database e controlla eventuali ripetizioni - OK
	Fa in modo che la proprietà "Duration" non sia di tipo String ma Timespan - OK
	Ricorda di esportare in CSV - OK
	A volte il nome file e la data di modifica non coincidono - OK

        Ho escluso tutti i file che cominciano con "Preview_" (per via dell'altro script)
#>

[System.Reflection.Assembly]::LoadWithPartialName("System.windows.Forms") | Out-Null
$directory = New-Object System.Windows.Forms.FolderBrowserDialog
$directory.RootFolder = "Desktop"; $directory.ShowNewFolderButton = $false

$directory.ShowDialog() | Out-Null

If ($directory.SelectedPath -eq ""){Write-Warning "No directory selected."; $directory.SelectedPath = ""; Exit}


$data = @()

#Acquisisco i nomi file
$temp = Get-ChildItem -Path $directory.SelectedPath -Recurse -File -Include ("*.mp4", "*.jpg") -Exclude "Preview_*" | Select-Object -Property Name
$temp = $temp.Name.Replace("-", "") #Rende simili i formati VID_ e QVR_
#Registro i nomi file nella proprietà NomiFile
$temp | ForEach-Object {$data += New-Object -TypeName psobject -Property @{NomeFile = $_}}

$temp = $null #Ripulisco perché inutile
$NotSupported = @()

$data | ForEach-Object {#Acquisisco e registro la data di acquisizione
    Write-Progress -Activity "Acquisendo e registrando dati nella proprietà 'TimeFromName'..."
    $time = $_.NomeFile

    If ($time.StartsWith("IMG_") -or $time.StartsWith("VID_")){
        $time = ($time[4..7] + "/" + $time[8,9] + "/" + $time[10,11] + "_" + $time[13,14] + ":" + $time[15,16] + ":" + $time[17,18]) #Array output es: 2 0 2 2 / 0 5 / 2 0 _ 1 3 : 3 4 : 2 1
        $time = $time -join ""
        $_ | Add-Member -MemberType NoteProperty -Name "TimeFromName" -Value ([datetime]::ParseExact($time, "yyyy/MM/dd_HH:mm:ss", $null))
    
    } ElseIf ($time.StartsWith("QVR_")){
        $time = ($time[4,5] + "/" + $time[6,7] + "/" + $time[8..11] + "_" + $time[13,14] + ":" + $time[15,16] + ":" + $time[17,18])
        $time = $time -join ""
        $_ | Add-Member -MemberType NoteProperty -Name "TimeFromName" -Value ([datetime]::ParseExact($time, "dd/MM/yyyy_HH:mm:ss", $null))
	    
    } ElseIf ($time.StartsWith("WIN_")) {
        $time = ($time[4..7] + "/" + $time[8,9] + "/" + $time[10,11] + "_" + $time[13,14] + ":" + $time[16,17] + ":" + $time[19,20])
        $time = $time -join ""
        $_ | Add-Member -MemberType NoteProperty -Name "TimeFromName" -Value ([datetime]::ParseExact($time, "yyyy/MM/dd_HH:mm:ss", $null))      
	
    } Else {
        $NotSupported += $time
        $_ | Add-Member -MemberType NoteProperty -Name "TimeFromName" -Value "Struttura non riconosciuta."
    }
}

$Length = @()
Get-ChildItem -Path $directory.SelectedPath -Recurse -File -Include ("*.mp4", "*.jpg") -Exclude "Preview_*"| ForEach-Object {

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

$total = $total | Format-List -Property "TotalDays", "TotalHours", "TotalMinutes", "TotalSeconds" 

Write-Host "Totale tempo registrato:" -ForegroundColor Green
$total

$data | Out-GridView -Title 'Output array $data'

If ( (Read-Host -Prompt "Esportare in CSV? [Y/N]").ToLower() -eq "y" ){$data | Export-Csv -Path $env:USERPROFILE\Desktop\Export.csv}
