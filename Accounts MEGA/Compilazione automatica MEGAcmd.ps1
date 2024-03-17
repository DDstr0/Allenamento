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

Write-Host "Avvio tra 5 secondi..."
Start-Sleep -Seconds 5

#INSERIMENTO DA MEGACMD CON LOGOUT AUTOMATICO
    For ($i = 0; $i -le $Emails.Length - 1; $i++){

        Start-Sleep -Milliseconds 10
        Set-Clipboard -Value ('login ' + $Emails[$i] + ' ' + $Passwords[$i])
        $wshell.SendKeys("^{v}") #CTRL + V
        $wshell.SendKeys('{ENTER}')
        Write-Warning ("Loggato in '"+$Emails[$i] + "'")
        Write-Host ""

        $wshell.SendKeys('logout')
        $wshell.SendKeys('{ENTER}')

        Start-Sleep -Seconds 3
    }
