#Inserisci progress bar

#L'idea sarebbe quella di usare lo stesso API che ho usato nel COD Project, in modo da poter usare lo script con la finestra diminuita
#Per dare la possibilità di scegliere i tasti da usare, creo 2 array paralleli che contengono i codici virtuali dei tasti (tipo 0x30) e i caratteri assegnati

$KeyChars = @("0", "1", "2", "3", "4", "5", "6", "7", "8", "9",
"a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z",
"Insert", "End", "DownArrow", "PageDown", "LeftArrow", "Clear", "RightArrow", "Home", "UpArrow", "PageUp",
"F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12")

$VirtualKeyCodes = @("0x30", "0x31", "0x32", "0x33", "0x34", "0x35", "0x36", "0x37", "0x38", "0x39",
"0x41", "0x42", "0x43", "0x44", "0x45", "0x46", "0x47", "0x48", "0x49", "0x4A", "0x4B", "0x4C", "0x4D", "0x4E", "0x4F", "0x50", "0x51", "0x52", "0x53", "0x54", "0x55", "0x56", "0x57", "0x58", "0x59", "0x5A",
"0x60", "0x61", "0x62", "0x63", "0x64", "0x65", "0x66", "0x67", "0x68", "0x69",
"0x70", "0x71", "0x72", "0x73", "0x74", "0x75", "0x76", "0x77", "0x78", "0x79", "0x7A", "0x7B")

#Attenzione all'uso delle variabili globali e locali
Function Select-Key (){
    Do {Write-Host "Configura il tasto da usare per avviare: " -NoNewLine
        $Selected = [Console]::ReadKey()
        Write-Host ""
        Write-Warning ("Il tasto attualmente selezionato è "+"'"+$Selected.Key+"'")
        Write-Host ""
        Write-Host "Procedere? [Y/N]: " -NoNewLine
        $Ans = Choice -c YN -n
        Write-Host ""; Write-Host ""
    } Until ($Ans -eq "Y" )

    $Global:SelectedChar = $Selected.Key

    #Controlla sia la proprietà .Key che .KeyChar
If ($KeyChars.IndexOf([string]$Selected.Key) -ne -1){
    $Global:Selected = $VirtualKeyCodes[$KeyChars.IndexOf([string]$Selected.Key)]
} Else {
    If ($KeyChars.IndexOf([string]$Selected.KeyChar) -ne -1){
    $Global:Selected = $VirtualKeyCodes[$KeyChars.IndexOf([string]$Selected.KeyChar)]
    } Else {
        Write-Warning "Il carattere non è presente."
        Write-Host ""
        Select-Key
    }
}

} Select-Key

[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.initialDirectory = "$env:USERPROFILE\Desktop"
$OpenFileDialog.Title = "Seleziona il file contenente gli indirizzi email (uno per riga)"
$OpenFileDialog.Filter = "Indirizzi Email (*.*)|*.*"

If ($OpenFileDialog.ShowDialog() -eq "Cancel"){Write-Warning "Nessun file selezionato."; Start-Sleep -Seconds 3; Exit}
$Emails = Get-Content -Path $OpenFileDialog.FileName

$OpenFileDialog.Title = "Seleziona il file contenente le password (una per riga)"
$OpenFileDialog.Filter = "Password (*.*)|*.*"
If ($OpenFileDialog.ShowDialog() -eq "Cancel"){Write-Warning "Nessun file selezionato."; Start-Sleep -Seconds 3; Exit}
$Passwords = Get-Content -Path $OpenFileDialog.FileName

If ($Emails.Length -ne $Passwords.Length){Write-Warning "I due file non contengono lo stesso numero di righe."; Start-Sleep -Seconds 3; Exit}


$wshell = New-Object -ComObject wscript.shell

$Signature = '[DllImport("user32.dll", CharSet=CharSet.Auto, ExactSpelling=true)] public static extern short GetAsyncKeyState(int virtualKeyCode);'
Add-Type -MemberDefinition $Signature -Name Keyboard -Namespace PsOneApi


Write-Host "Seleziona il metodo: "
Write-Host "[1] " -NoNewLine -ForegroundColor "Green"; Write-Host "Inserimento da browser"
Write-Host "[2] " -NoNewLine -ForegroundColor "Green"; Write-Host "Inserimento da MegaCMD"
Write-Host ""

#Write-Host "> " -NoNewLine
$Ans = Choice -c 12 -n
#Write-Host ""; Write-Host ""

#Note: tutti questi Start-Sleep servono perché altrimenti si mangia alcune operazioni
Switch ($Ans){
    1 {#INSERIMENTO DA BROWSER
    For ($i = 0; $i -le $Emails.Length - 1; $i++){
        $KeyPressed = $false
        Write-Warning ("Premi '" + $SelectedChar + "' per continuare")
        Do {
            If ( [bool]([PsOneApi.Keyboard]::GetAsyncKeyState($Selected) -eq -32767) ){$KeyPressed = $true}
            Start-Sleep -Milliseconds 50
        } Until ($KeyPressed)

        Start-Sleep -Milliseconds 10
        Set-Clipboard -Value ([string]$Emails[$i]) #Siccome il metodo SendKeys non manda alcuni caratteri come "." o "+"
        $wshell.SendKeys("^{v}") #CTRL + V
        Start-Sleep -Milliseconds 10
        $wshell.SendKeys('{TAB}')
        Start-Sleep -Milliseconds 10
        Set-Clipboard -Value $Passwords[$i]
        $wshell.SendKeys("^{v}") #CTRL + V
        Start-Sleep -Milliseconds 10
        $wshell.SendKeys('{ENTER}')

        Write-Warning ("Loggato in '"+$Emails[$i] + "'")
        Write-Host ""
    }
}
    2 {#INSERIMENTO DA MEGACMD CON LOGOUT AUTOMATICO
    For ($i = 0; $i -le $Emails.Length - 1; $i++){
        $KeyPressed = $false
        Write-Warning ("Premi '" + $SelectedChar + "' per loggare")
        Do {
            If ( [bool]([PsOneApi.Keyboard]::GetAsyncKeyState($Selected) -eq -32767) ){$KeyPressed = $true}
            Start-Sleep -Milliseconds 50
        } Until ($KeyPressed)

        Start-Sleep -Milliseconds 10
        Set-Clipboard -Value ('login ' + $Emails[$i] + ' ' + $Passwords[$i])
        $wshell.SendKeys("^{v}") #CTRL + V
        $wshell.SendKeys('{ENTER}')
        Write-Warning ("Loggato in '"+$Emails[$i] + "'")
        Write-Host ""

        $wshell.SendKeys('logout')
        $wshell.SendKeys('{ENTER}')
    }
}
}