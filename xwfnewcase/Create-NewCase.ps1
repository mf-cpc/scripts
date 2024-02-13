if (-not ([System.Management.Automation.PSTypeName]'Console.Window').Type) {
    Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
    $consolePtr = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($consolePtr, 0)
}
Add-Type -Assembly System.Windows.Forms

function Show-NewCase {

    # Set the label prompt and title prompt if not specified when calling the function
    param (
        [string]$Prompt = 'Case number / Numéro de dossier:',
        [string]$Title = 'New Case / Nouveau dossier'
    )

    $newCaseForm = New-Object Windows.Forms.Form
    $newCaseForm.Text = $Title
    $newCaseForm.Size = New-Object Drawing.Size @(225, 140)
    $newCaseForm.StartPosition = 'CenterScreen'
    $newCaseForm.FormBorderStyle = 'FixedDialog'
    $newCaseForm.MaximizeBox = $false
    $newCaseForm.MinimizeBox = $false

    $label = New-Object Windows.Forms.Label
    $label.Location = New-Object Drawing.Point @(10, 20)
    $label.Size = New-Object Drawing.Size @(210, 20)
    $label.Text = $Prompt

    $textBox = New-Object Windows.Forms.TextBox
    $textBox.Location = New-Object Drawing.Point @(10, 40)
    $textBox.Size = New-Object Drawing.Size @(195, 20)
    $textBox.Text = (Get-Date).Year.ToString() + '-'
    $textBox.Select($textBox.Text.Length, 0)

    $buttonOK = New-Object Windows.Forms.Button
    $buttonOK.Location = New-Object Drawing.Point @(10, 70)
    $buttonOK.Size = New-Object Drawing.Size @(75, 23)
    $buttonOK.Text = 'OK'
    $buttonOK.TextAlign = [System.Drawing.ContentAlignment]::BottomCenter
    $buttonOK.DialogResult = [Windows.Forms.DialogResult]::OK
    $buttonOK.Add_Click({ $newCaseForm.close() })

    $buttonCancel = New-Object Windows.Forms.Button
    $buttonCancel.Location = New-Object Drawing.Point @(90, 70)
    $buttonCancel.Size = New-Object Drawing.Size @(75, 23)
    $buttonCancel.Text = 'Cancel'
    $buttonCancel.TextAlign = [System.Drawing.ContentAlignment]::BottomCenter
    $buttonCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $newCaseForm.Controls.Add($label)
    $newCaseForm.Controls.Add($textBox)
    $newCaseForm.Controls.Add($buttonOK)
    $newCaseForm.Controls.Add($buttonCancel)

    $newCaseForm.AcceptButton = $buttonOK
    $newCaseForm.TopMost = $true
    
    $result = $newCaseForm.ShowDialog()

    if ($result -eq [Windows.Forms.DialogResult]::OK -and $textBox.Text -ne '') {
        $textBox.Text
    }
    else {
        $null
    }
}

$folderName = Show-NewCase

if ($null -ne $folderName) {
    try {
        # Set source template folder and destination folder paths
        $templateFolder = 'C:\CASE_FOLDER_STRUCTURE'
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
            Write-Progress -Activity 'Copying Files' -Status "Progress: $percentage%" -PercentComplete $percentage
        }

        # Create a shortcut on the desktop for xf64.exe in the xwf folder of the copied folder
        $shortcutPath = [System.IO.Path]::Combine($env:USERPROFILE, 'Desktop', "$folderName.lnk")
        $programPath = [System.IO.Path]::Combine($destinationPath, 'xwf', 'xwforensics64.exe')
        if (Test-Path $programPath) {
            # Create the shortcut
            $WScriptShell = New-Object -ComObject wscript.shell
            $Shortcut = $WScriptShell.CreateShortcut($shortcutPath)
            $Shortcut.TargetPath = $programPath
            $Shortcut.Save()
            $shortcutBytes = [System.IO.File]::ReadAllBytes($shortcutPath)
            $shortcutBytes[0x15] = $shortcutBytes[0x15] -bor 0x20
            [System.IO.File]::WriteAllBytes($shortcutPath, $shortcutBytes)
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("The file $programPath does not exist!", 'Error', 'OK', 'Error')
            return
        }
        [System.Windows.Forms.MessageBox]::Show("New case folder created at $destinationPath with relevant shortcut", 'Success', 'OK', 'Information')
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Unable to complete: $_.ScriptStackTrace", 'Error', 'OK', 'Error')
    }
}