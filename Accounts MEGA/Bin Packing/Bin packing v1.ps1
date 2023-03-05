#$ErrorActionPreference = "SilentlyContinue" ### PRESTA ATTENZIONE NEL DEBUGGING ###

[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$FolderBrowser.Description = "Seleziona la cartella contenente gli elementi da ordinare" #Le cartelle presenti all'interno verranno trattate come singoli file
$FolderBrowser.ShowNewFolderButton = $False

If ($FolderBrowser.ShowDialog() -eq "Cancel"){Write-Warning "Nessuna cartella selezionata"; Exit}

Write-Host "Seleziona l'unità  di misura: "
Write-Host "[1] " -ForegroundColor "Green" -NoNewLine; Write-Host "GB"
Write-Host "[2] " -ForegroundColor "Green" -NoNewLine; Write-Host "MB"
Write-Host ""
Write-Host "> " -NoNewLine; $Unit = Choice /c 12 /n

Switch($Unit){
    1 {$Unit = 'GB'}
    2 {$Unit = 'MB'}
} Write-Host "$Unit"

Write-Host ""

Write-Host "Inserisci la dimensione del contenitore: "
Write-Host ""

Do {Write-Host "> " -NoNewLine; [float]$Storage = Read-Host} Until ([float]$Storage -gt 0)
Write-Host ""

$AvailableSpace = [string]$Storage+$Unit #Risultato espresso in Bytes


[System.Collections.ArrayList]$Items = @() #Il tipo ArrayList è importante per permettere la rimozione tramite medoto Remove()
Get-ChildItem -Path $FolderBrowser.SelectedPath <#| Where-Object -Property "Length" -LE $AvailableSpace #> | ForEach-Object { $Items += New-Object -TypeName PSObject -Property @{FullName = $_.FullName; Size = [Int64]$_.Length; IsContainer = $_.PSIsContainer} }
$Count = $Items.FullName #Archivio il nome completo di tutti i file della cartella


#Sapendo che i PSObject (e forse le HashTable) riflettono i valori e sono anti ridondanza, se eseguo una modifica ad un oggetto di $Folders, la modifica arriva anche all'elemento di $Items
$Folders = $Items | Where-Object -Property "IsContainer" -EQ $True

For ($i = 0; $i -le $Folders.Count - 1; $i++){
    [Int64]$Temp = 0
    Get-ChildItem -Path $Folders[$i].FullName -Recurse | ForEach-Object {$Temp += $_.Length}
    #Write-Warning ("Modifica: " + $Folders[$i].Size + " --> $Temp")
    $Folders[$i].Size = $Temp
} $Folders = $null

$Items = $Items | Where-Object -Property "Size" -LE $AvailableSpace | Sort-Object -Property "Size" -Descending
$Count += $Items.FullName #Dopo aver rimosso tutti i file troppo pesanti, unisco la lista completa vecchia a quella attuale

$Count = ($Count | Group-Object | Where-Object -Property "Count" -EQ 1).Name #Conto quali file sono presenti una sola volta

	#UN ALTRO METODO POTEVA ESSERE ARCHIVIARE SEMPLICEMENTE NEL PSOBJECT ANCHE IL NOME E SCRIVERE '$Count = $Items.Name'
#For ($i = 0; $i -le $Count.Length - 1; $i++){
	#         ($Count[0])[40..50] -Join "" #Serve a mostrarmi solo quello che viene dopo l'ultimo "\"
#    $Count[$i] = ($Count[$i])[ ($Count[$i].LastIndexOf("\") + 1) .. ($Count[$i].Length - 1) ] -join ""
#}


If ($Items.Count -eq 0){
    Write-Warning "Nessun elemento della cartella selezionata rispetta i criteri inseriti."
    Exit
}

If ($Count.Length -gt 0){
    Write-Warning "I seguenti elementi occupano troppo spazio, pertanto sono stati esclusi dall'operazione: "
    Write-Host ""    

    For ($i = 0; $i -le $Count.Length - 1; $i++){
        Write-Host ( '[' + [string]($i + 1) + ']' + "`t" + $Count[$i] )
    }

    Write-Host ""
}

[string]$FolderName = Get-Date -Format "[HH.mm.ss]"
New-Item -Path $FolderBrowser.SelectedPath -ItemType "Directory" -Name $FolderName | Out-Null
Write-Warning ("Nuova cartella '" + $FolderName + "' creata.")


Do {
    Move-Item -Path $Items[0].FullName -Destination ($FolderBrowser.SelectedPath+'\'+$FolderName)
    Write-Warning ("Spostato elemento '" + $Items[0].FullName + "'")
    $AvailableSpace -= $Items[0].Size
    $Items.RemoveAt(0)

    #Genera errore (ma funziona lo stesso) perché per aver rimosso il primo elemento ed essere arrivato a .Count = 0, non posso porre uguale l'ArrayList a qualcosa che non c'è (forse)

    #$Items | Where-Object -Property "Size" -LE $AvailableSpace | Sort-Object -Property "Size" -Descending; Pause

    $Items = @($Items | Where-Object -Property "Size" -LE $AvailableSpace | Sort-Object -Property "Size" -Descending)
} Until ($Items.Count -eq 0)

Write-Host ""; Write-Warning ( "Spazio rimanente: " + [Math]::Round($AvailableSpace/1MB, 2) + " MB")

#$ErrorActionPreference = "Continue"