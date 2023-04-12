Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Define the GUI form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Furigana remover"
$form.Size = New-Object System.Drawing.Size(400,200)
$form.StartPosition = "CenterScreen"

# Define the file input control
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = "Select an .epub file to process:"
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,40)
$textBox.Size = New-Object System.Drawing.Size(280,20)
$form.Controls.Add($textBox)

$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(300,40)
$button.Size = New-Object System.Drawing.Size(80,20)
$button.Text = "Browse..."
$button.Add_Click({
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Filter = "EPUB files (*.epub)|*.epub"
    $fileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    $result = $fileDialog.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $textBox.Text = $fileDialog.FileName
    }
})
$form.Controls.Add($button)

# Define the backup checkbox
$backupCheckbox = New-Object System.Windows.Forms.CheckBox
$backupCheckbox.Location = New-Object System.Drawing.Point(10,100)
$backupCheckbox.Size = New-Object System.Drawing.Size(120,30)
$backupCheckbox.Text = "Take backup"
$form.Controls.Add($backupCheckbox)

# Define the execute button
$executeButton = New-Object System.Windows.Forms.Button
$executeButton.Location = New-Object System.Drawing.Point(10,70)
$executeButton.Size = New-Object System.Drawing.Size(120,30)
$executeButton.Text = "Remove furigana"
$executeButton.Add_Click({
    $filePath = $textBox.Text
    if ([System.IO.Path]::GetExtension($filePath) -eq ".epub" -and [System.IO.File]::Exists($filePath)) {
        if ($backupCheckbox.Checked) {
            $backupFilePath = $filePath -replace "\.epub$", "_backup.epub"
            Copy-Item -Path $filePath -Destination $backupFilePath
        }
        $epubPath = $filePath
        # Load the System.IO.Compression.FileSystem assembly
        Add-Type -AssemblyName System.IO.Compression.FileSystem

        # Extract the contents of the ePub file to a temporary directory
        $epubDir = [System.IO.Path]::GetFullPath("$env:TEMP\epub-$([System.Guid]::NewGuid().ToString())")
        [System.IO.Compression.ZipFile]::ExtractToDirectory($epubPath, $epubDir)

        # Delete the furigana elements from all HTML files in the ePub directory
        $htmlFiles = Get-ChildItem -Path $epubDir -Recurse -Include "*.xhtml"
        foreach ($file in $htmlFiles) {
            $content = Get-Content $file.FullName
            $content = $content -replace "<ruby>.*?</ruby>", ""
            Set-Content -Path $file.FullName -Value $content
        }

        # Re-zip the contents of the ePub directory
        $zipPath = [System.IO.Path]::GetFullPath("$epubPath.tmp")
        [System.IO.Compression.ZipFile]::CreateFromDirectory($epubDir, $zipPath)

        # Replace the original ePub file with the new one
        Remove-Item $epubPath
        Rename-Item $zipPath -NewName ([System.IO.Path]::GetFileName($epubPath))
        Write-Host "Furigana removed successfully"
        [System.Windows.Forms.MessageBox]::Show("Furigana removed successfully")
    } elseif ([System.IO.Path]::GetExtension($filePath) -ne ".epub") {
        [System.Windows.Forms.MessageBox]::Show("File must be an .epub file")
    } else {
        [System.Windows.Forms.MessageBox]::Show("File does not exist")
    }
})
$form.Controls.Add($executeButton)

# Display the GUI form
$form.ShowDialog() | Out-Null
