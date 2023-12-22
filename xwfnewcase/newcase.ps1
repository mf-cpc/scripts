$destinationFolder = Read-Host "Enter case number"

# Set source template folder and destination folder paths
$templateFolder = "C:\CASE_FOLDER_STRUCTURE"
$destinationPath = "C:\$destinationFolder"

# Copy the template folder to the destination folder
Copy-Item -Path $templateFolder -Destination $destinationPath -Recurse

# Create a shortcut on the desktop for xf64.exe in the xwf folder of the copied folder
$shortcutPath = [System.IO.Path]::Combine($env:USERPROFILE, 'Desktop', '$destinationFolder.lnk')
$programPath = [System.IO.Path]::Combine($destinationPath,'xwf','xwforensics64.exe')

# Create the shortcut
$WScriptShell = New-Object -ComObject wscript.shell
$Shortcut = $WScriptShell.CreateShortcut($shortcutPath)
$Shortcut.TargetPath = $programPath
$Shortcut.Save()

Write-Host "Created the new case and shortcut."
