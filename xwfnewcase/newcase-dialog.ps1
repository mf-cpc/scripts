Add-Type -Assembly System.Windows.Forms

function Show-NewCase {

    # Set the label prompt and title prompt if not specified when calling the function
    param (
        [string]$Prompt = "Enter the new case number:",
        [string]$Title = "Create new Case"
    )

    $newCaseForm = New-Object Windows.Forms.Form
    $newCaseForm.Text = $Title
    $newCaseForm.Size = New-Object Drawing.Size @(300,150)
    $newCaseForm.StartPosition = 'CenterScreen'

    $label = New-Object Windows.Forms.Label
    $label.Location = New-Object Drawing.Point @(10,20)
    $label.Size = New-Object Drawing.Size @(280,20)
    $label.Text = $Prompt

    $textBox = New-Object Windows.Forms.TextBox
    $textBox.Location = New-Object Drawing.Point @(10,40)
    $textBox.Size = New-Object Drawing.Size @(260,20)
    $textBox.Text = (Get-Date).Year.ToString() + "-"
    $textBox.Select($textBox.Text.Length, 0)

    $buttonOK = New-Object Windows.Forms.Button
    $buttonOK.Location = New-Object Drawing.Point @(10,70)
    $buttonOK.Size = New-Object Drawing.Size @(75,23)
    $buttonOK.Text = "OK"
    $buttonOK.DialogResult = [Windows.Forms.DialogResult]::OK
    $buttonOK.Add_Click({ $newCaseForm.close() })

    $buttonCancel = New-Object Windows.Forms.Button
    $buttonCancel.Location = New-Object Drawing.Point @(90,70)
    $buttonCancel.Size = New-Object Drawing.Size @(75,23)
    $buttonCancel.Text = "Cancel"
    $buttonCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $newCaseForm.Controls.Add($label)
    $newCaseForm.Controls.Add($textBox)
    $newCaseForm.Controls.Add($buttonOK)
    $newCaseForm.Controls.Add($buttonCancel)

    $newCaseForm.AcceptButton = $buttonOK
    $newCaseForm.TopMost = $true
    
    $result = $newCaseForm.ShowDialog()

    if ($result -eq [Windows.Forms.DialogResult]::OK -and $textBox.Text -ne "") {
        $textBox.Text
    } else {
        $null
    }
}

# $folderName = Show-NewCase -Prompt "Numéro de dossier:" -Title "Nouveau dossier"
$folderName = Show-NewCase

if ($null -ne $folderName) {

    # Set source template folder and destination folder paths
    $templateFolder = "C:\CASE_FOLDER_STRUCTURE"
    $destinationPath = "C:\$folderName"

    # Progress bar calculation
    $totalSize = (Get-ChildItem -Path $templateFolder -Recurse | Measure-Object -Property Length -Sum).Sum
    $copiedSize = 0


    # Copy the template folder to the destination folder with progress bar
    Get-ChildItem -Path $templateFolder -Recurse | ForEach-Object {
        $destinationFile = $_.FullName -replace [regex]::Escape($templateFolder), $destinationPath
        Copy-Item -Path $_.FullName -Destination $destinationFile -Force
        $copiedSize += $_.Length
        $percentage = ($copiedSize / $totalSize) * 100
        Write-Progress -Activity "Copying Files" -Status "Progress: $percentage%" -PercentComplete $percentage
    }
    # Copy-Item -Path $templateFolder -Destination $destinationPath -Recurse

    # Create a shortcut on the desktop for xf64.exe in the xwf folder of the copied folder
    $shortcutPath = [System.IO.Path]::Combine($env:USERPROFILE, 'Desktop', "$folderName.lnk")
    $programPath = [System.IO.Path]::Combine($destinationPath,'xwf','xwforensics64.exe')

    # Create the shortcut
    $WScriptShell = New-Object -ComObject wscript.shell
    $Shortcut = $WScriptShell.CreateShortcut($shortcutPath)
    $Shortcut.TargetPath = $programPath
    $Shortcut.Save()

    [System.Windows.Forms.MessageBox]::Show("New case folder created at $destinationPath and relevant shortcut", "Success", 'OK', 'Information')

    # Write-Host "New case number is $folderName"
}
# use for debuggin
# else {
#     Write-Host "Operation canceled"
#     exit
# }
