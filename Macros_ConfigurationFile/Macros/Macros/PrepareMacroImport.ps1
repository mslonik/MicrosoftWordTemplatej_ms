Add-Type -AssemblyName System.Windows.Forms
Start-Sleep -Seconds 1

# Function to show the folder selection dialog
function Select-FolderDialog {
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select the input folder"
    $folderBrowser.ShowNewFolderButton = $true
    $result = $folderBrowser.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderBrowser.SelectedPath
    } else {
        Write-Host "No folder selected. Exiting script."
        exit
    }
}

# Function to get the folder path from the user
function Get-FolderPath {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Input Folder"
    $form.Width = 500
    $form.Height = 150

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Select the input folder or enter the path:"
    $label.AutoSize = $true
    $label.Top = 20
    $label.Left = 10
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Width = 400
    $textBox.Top = 50
    $textBox.Left = 10
    $form.Controls.Add($textBox)

    $buttonBrowse = New-Object System.Windows.Forms.Button
    $buttonBrowse.Text = "Browse"
    $buttonBrowse.Top = 50
    $buttonBrowse.Left = 420
    $buttonBrowse.Add_Click({
        $selectedPath = Select-FolderDialog
        if ($selectedPath) {
            $textBox.Text = $selectedPath
        }
    })
    $form.Controls.Add($buttonBrowse)

    $buttonSubmit = New-Object System.Windows.Forms.Button
    $buttonSubmit.Text = "Submit"
    $buttonSubmit.Top = 80
    $buttonSubmit.Left = 10
    $buttonSubmit.Add_Click({
        if ($textBox.Text -and (Test-Path -Path $textBox.Text) -and (Get-Item -Path $textBox.Text).PSIsContainer) {
            $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.Close()
        } else {
            [System.Windows.Forms.MessageBox]::Show("Invalid path entered. Exiting script.")
            $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $form.Close()
        }
    })
    $form.Controls.Add($buttonSubmit)

    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $textBox.Text
    } else {
        Write-Host "No valid path entered. Exiting script."
        exit
    }
}

# Get the input folder from the user
$inputFolder = Get-FolderPath

# Define the file types to search for
$fileTypes = @("*.bas", "*.cls", "*.frm", "*.frx")

# Initialize an empty array to store the list of files
$filesToCopy = @()

# Loop through each file type and search for files
foreach ($fileType in $fileTypes) {
    $files = Get-ChildItem -Path $inputFolder -Recurse -Filter $fileType
    $filesToCopy += $files.FullName
}

# Check if the list of files is empty
if ($filesToCopy.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show("No files found to copy. Exiting script.")
    exit
}

# Function to get the output folder path from the user
function Get-OutputFolderPath {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Specify the Output Folder"
    $form.Width = 500
    $form.Height = 150

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Select the output folder or enter the path:"
    $label.AutoSize = $true
    $label.Top = 20
    $label.Left = 10
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Width = 400
    $textBox.Top = 50
    $textBox.Left = 10
    $form.Controls.Add($textBox)

    $buttonBrowse = New-Object System.Windows.Forms.Button
    $buttonBrowse.Text = "Browse"
    $buttonBrowse.Top = 50
    $buttonBrowse.Left = 420
    $buttonBrowse.Add_Click({
        $selectedPath = Select-FolderDialog
        if ($selectedPath) {
            $textBox.Text = $selectedPath
        }
    })
    $form.Controls.Add($buttonBrowse)

    $buttonSubmit = New-Object System.Windows.Forms.Button
    $buttonSubmit.Text = "Submit"
    $buttonSubmit.Top = 80
    $buttonSubmit.Left = 10
    $buttonSubmit.Add_Click({
        if ($textBox.Text -and (Test-Path -Path $textBox.Text) -and (Get-Item -Path $textBox.Text).PSIsContainer) {
            $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.Close()
        } else {
            [System.Windows.Forms.MessageBox]::Show("Invalid path entered. Exiting script.")
            $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $form.Close()
        }
    })
    $form.Controls.Add($buttonSubmit)

    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $textBox.Text
    } else {
        Write-Host "No valid path entered. Exiting script."
        exit
    }
}

# Get the output folder from the user
$outputFolder = Get-OutputFolderPath

# Copy each file to the output folder with error checking
foreach ($file in $filesToCopy) {
    try {
        $destinationPath = Join-Path -Path $outputFolder -ChildPath (Split-Path -Leaf $file)
        Copy-Item -Path $file -Destination $destinationPath -Force -ErrorAction Stop
        Start-Sleep -Milliseconds 100
        if (Test-Path -Path $destinationPath) {
            # Successfully copied
        } else {
            throw "Failed to copy: $file"
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to copy: $file. Error: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        exit
    }
}

# Show final message box with information icon and play a sound
[System.Windows.Forms.MessageBox]::Show("File copy operation completed successfully.`nDestination path: $outputFolder", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
[System.Media.SystemSounds]::Asterisk.Play()