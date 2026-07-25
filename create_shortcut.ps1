$desktop = [Environment]::GetFolderPath("Desktop")

# Remove old shortcuts
Get-ChildItem -Path $desktop | Where-Object { $_.Name -like "*ForgePDF*" -or $_.Name -like "*PDF_Converter*" } | Remove-Item -Force -ErrorAction SilentlyContinue

# Create WScript Shell shortcut
$wsh = New-Object -ComObject WScript.Shell
$sc = $wsh.CreateShortcut("$desktop\ForgePDF.lnk")
$sc.TargetPath = "C:\Windows\System32\wscript.exe"
$sc.Arguments = '"C:\Users\uSer\ForgePDF\ForgePDF.vbs"'
$sc.WorkingDirectory = "C:\Users\uSer\ForgePDF"
$sc.Save()

Write-Host "ForgePDF Desktop Shortcut created successfully!"
