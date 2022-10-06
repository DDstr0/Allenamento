[System.Reflection.Assembly]::LoadWithPartialName("System.windows.Forms") | Out-Null
$directory = New-Object System.Windows.Forms.FolderBrowserDialog
$directory.RootFolder = "Desktop"; $directory.ShowNewFolderButton = $false

$directory.ShowDialog() | Out-Null

Get-ChildItem -Path $directory.SelectedPath -Directory <#| Where-Object -Property Length -EQ 10 #> | ForEach-Object {$date = [datetime]::Parse($_.Name); Set-ItemProperty -Path $_.FullName -Name "LastWriteTime" -Value $date}
