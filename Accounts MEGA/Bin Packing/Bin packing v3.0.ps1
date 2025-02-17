Write-Host "Debug mode? [Y/N]: " -NoNewline -ForegroundColor Green
If ((Choice -c YN -n) -eq "Y"){Write-Host "Y"; $DebugMode = @{WhatIf = $True}} Else {Write-Host "N"; $DebugMode = @{WhatIf = $False}}

[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$FolderBrowser.Description = "Seleziona la cartella contenente gli elementi da ordinare" #Le cartelle presenti all'interno verranno trattate come singoli file
$FolderBrowser.ShowNewFolderButton = $False

If ($FolderBrowser.ShowDialog() -eq "Cancel"){Write-Warning "Nessuna cartella selezionata"; Exit}

Write-Host ""
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


If ( (Test-Path -Path ($FolderBrowser.SelectedPath+'\Packs') ) -eq $False){
    New-Item -Path $FolderBrowser.SelectedPath -Name "Packs" -ItemType "Directory" @DebugMode | Out-Null
    Write-Warning ("Cartella 'Packs' creata.")
} Else {
    Write-Warning ("Cartella 'Packs' già esistente.")
}

$Items = @()
Get-ChildItem -Path $FolderBrowser.SelectedPath -Exclude "Packs" | ForEach-Object { $Items += New-Object -TypeName PSObject -Property @{FullName = $_.FullName; Size = [Int64]$_.Length; IsContainer = $_.PSIsContainer} }
$Count = @($Items.FullName) #Archivio il nome completo di tutti i file della cartella


$Folders = @($Items | Where-Object -Property "IsContainer" -EQ $True)

#Sapendo che i PSObject (e forse le HashTable) riflettono i valori e sono anti ridondanza, se eseguo una modifica ad un oggetto di $Folders, la modifica arriva anche all'elemento di $Items
For ($i = 0; $i -le $Folders.Count - 1; $i++){
    [Int64]$Temp = 0
    Get-ChildItem -Path $Folders[$i].FullName -Recurse | ForEach-Object {$Temp += $_.Length}
    #Write-Warning ("Modifica: " + $Folders[$i].Size + " --> $Temp")
    $Folders[$i].Size = $Temp
} $Folders = $null


$Items = @($Items | Where-Object -Property "Size" -LE $AvailableSpace | Sort-Object -Property "Size" -Descending)
$Count += @($Items.FullName) #Dopo aver rimosso tutti i file troppo pesanti, unisco la lista completa vecchia a quella attuale

$Count = @($Count | Group-Object | Where-Object -Property "Count" -EQ 1).Name #Conto quali file sono presenti una sola volta

	#UN ALTRO METODO POTEVA ESSERE ARCHIVIARE SEMPLICEMENTE NEL PSOBJECT ANCHE IL NOME E SCRIVERE '$Count = $Items.Name'
#For ($i = 0; $i -le $Count.Length - 1; $i++){
	#         ($Count[0])[40..50] -Join "" #Serve a mostrarmi solo quello che viene dopo l'ultimo "\"
#    $Count[$i] = ($Count[$i])[ ($Count[$i].LastIndexOf("\") + 1) .. ($Count[$i].Length - 1) ] -join ""
#}


If ($Count.Length -gt 0){
    Write-Warning "I seguenti elementi superano i $Storage $Unit, pertanto sono stati esclusi dall'operazione: "
    Write-Host ""    

    For ($i = 0; $i -le $Count.Length - 1; $i++){
        Write-Host ( '[' + [string]($i + 1) + ']' + "`t" + $Count[$i] )
    }

    Write-Host ""

} ElseIf ($Count.Length -eq $Items.Count) {
    Write-Warning "Nessun elemento della cartella selezionata rispetta i criteri inseriti."
    Exit
}


$FolderName = Get-Date -Format "yyyy.MM.dd_HH.mm.ss"
New-Item -Path ($FolderBrowser.SelectedPath+'\Packs') -ItemType "Directory" -Name $FolderName @DebugMode | Out-Null
Write-Warning ("Nuovo pack '" + $FolderName + "' creato.")


    #$Items và forzato ad essere un array, altrimenti l'attributo Count sparisce quando rimane un solo elemento e il ciclo esce senza spostare l'ultimo elemento (lasciando spazio occupabile)
Do {$Items = @($Items | Where-Object -Property "Size" -LE $AvailableSpace | Sort-Object -Property "Size" -Descending)

    If ($Items.Count -gt 0){ #Altrimenti finisco oltre i limiti della matrice
        Move-Item -Path $Items[0].FullName -Destination ($FolderBrowser.SelectedPath+'\Packs\'+$FolderName) @DebugMode | Out-Null
        Write-Warning ("Spostato elemento '" + $Items[0].FullName + "' (" + [Math]::Round($Items[0].Size/1MB, 2) + " MB)")

        $AvailableSpace -= $Items[0].Size
        $Items[0] = $Null
    }

} Until ($Items.Count -eq 0)

Write-Host ""

If ($AvailableSpace/1MB -ge 1024){
    Write-Warning ("Spazio rimanente: " + [Math]::Round($AvailableSpace/1GB, 2) + " GB su " + "$Storage $Unit (occupato al " + [string](100 - [Math]::Round($AvailableSpace/([string]$Storage+$Unit) * 100, 2)) + " %)")
} Else {
    Write-Warning ("Spazio rimanente: " + [Math]::Round($AvailableSpace/1MB, 2) + " MB su " + "$Storage $Unit (occupato al " + [string](100 - [Math]::Round($AvailableSpace/([string]$Storage+$Unit) * 100, 2)) + " %)")
}
