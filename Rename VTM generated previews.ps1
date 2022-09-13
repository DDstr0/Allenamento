#Obiettivo: per ciascun file mp4 controlla se esiste un file jpg che contiene il nome del video e aggiungi il prefisso o suffisso "preview"
#Implementazioni: Log per ripristino
#Correzioni bug/migliorie: Il parametro "NewName" del cmdlet "Rename-Item" non richiede necessariamente l'intero percorso file; rimuovi alcune variabili inutili e inserisci solo i nomi; nei log invece è necessario che rimangano

$OutputEncoding = [System.Text.Encoding]::UTF8 #Codifica degli output (per evitare problemi con gli accenti)

Write-Warning ("I log saranno generati sul Desktop di '" + $env:USERNAME + "'"); Write-Host ""

If ( (Read-Host -Prompt "DISATTIVARE protezione modifiche WhatIf? [Y/N]").ToLower() -eq "y"){
    $Protection = @{WhatIf = $false}
    Write-Host "Risposta affermativa esplicita. PROTEZIONE DISATTIVA." -ForegroundColor "Red"
    
} Else {
    $Protection = @{WhatIf = $true}
    Write-Host "Risposta negativa o non esplicita. PROTEZIONE ATTIVA." -ForegroundColor "Green"
} Write-Host ""


[System.Reflection.Assembly]::LoadWithPartialName("System.windows.Forms") | Out-Null
$directory = New-Object System.Windows.Forms.FolderBrowserDialog
$directory.RootFolder = "Desktop"; $directory.ShowNewFolderButton = $false

If ($directory.ShowDialog() -eq "Cancel"){Write-Warning "Nessuna cartella selezionata."; Exit} #La chiamata al metodo ShowDialog viene eseguita sempre e comunque


If ( (Read-Host -Prompt "Eliminare i file '.vtx'? [Y/N]").ToLower() -eq "y"){
    $temp = (Get-ChildItem -Path $directory.SelectedPath -File -Recurse -Include "*.vtx").Count
    Get-ChildItem -Path $directory.SelectedPath -File -Recurse -Include "*.vtx" | Remove-Item
    Write-Host "Rimossi $temp file .vtx" -ForegroundColor "Yellow"
} Else {Write-Host "Rimossi 0 file .vtx" -ForegroundColor "Green"}
Write-Host ""


Write-Host "Inizio procedura... [Premere INVIO per iniziare o CTRL+C per uscire]: " -NoNewLine; Read-Host
Write-Host ""


$FullNames = 	(Get-ChildItem -Path $directory.SelectedPath -File -Recurse -Include "*.mp4").FullName
$Names = 	(Get-ChildItem -Path $directory.SelectedPath -File -Recurse -Include "*.mp4").Name		#NEW
$Directories = 	(Get-ChildItem -Path $directory.SelectedPath -File -Recurse -Include "*.mp4").DirectoryName	#NEW
$NoExt = 	$Names.Trim('.mp4')	#Lascia i nomi dei file senza estensioni
$FullNoExt = 	$FullNames.Trim('.mp4') #Lascia i nomi dei file senza estensioni con il percorso		#NEW

$LogPath = ("$env:USERPROFILE\Desktop" + "\Log di ripristino " + (Get-Date -Format "dd-MM-yyyy_HH-mm-ss") + ".log")

For ($i = 0; $i -le $FullNames.Count - 1; $i++){
    Write-Progress -Activity "Controllando e rinominando le preview grid..." -Status ([string]([Math]::Round(($i / $FullNames.Count), 2) * 100) + "% completato.") -PercentComplete (($i / $FullNames.Count) * 100)

    $temp = ($FullNames[$i] + ".jpg") #I nomi file delle preview di Video Thumbnails Maker sono *Nome video* + ".jpg"

    If ( [System.IO.File]::Exists("$temp") ){
        #Rename-Item -Path "$temp" -NewName ($FullNoExt[$i] + "_preview.jpg") @Protection 				#Suffisso
        Rename-Item -Path "$temp" -NewName ($Directories[$i] + "\" + "Preview_" + $NoExt[$i] + ".jpg") @Protection 	#Prefisso
        Write-Host ("Rinominata preview di '" + $FullNames[$i] + "'.") -ForegroundColor "Green"

        #Write-Output ($temp + " --> " + $FullNoExt[$i] + "_preview.jpg") | Out-File -FilePath "$LogPath" -Append				#Suffisso
        Write-Output ($temp + " --> " + $Directories[$i] + "\" + "Preview_" + $NoExt[$i] + ".jpg") | Out-File -FilePath "$LogPath" -Append	#Prefisso
    } Else {
        Write-Host ("Per '" + $FullNames[$i] + "' non è stata generata una preview.") -ForegroundColor "Red"
    }
}