Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to create and handle the message box with options
function Show-CompletionDialog($outputFile) {
    $msgBox = New-Object System.Windows.Forms.Form
    $msgBox.Text = 'Task Completed'
    $msgBox.Size = New-Object System.Drawing.Size(300, 200)
    $msgBox.StartPosition = 'CenterScreen'

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Select an option:"
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(280, 20)
    $msgBox.Controls.Add($label)

    $buttonOpenFile = New-Object System.Windows.Forms.Button
    $buttonOpenFile.Text = "Open File"
    $buttonOpenFile.Location = New-Object System.Drawing.Point(10, 50)
    $buttonOpenFile.Size = New-Object System.Drawing.Size(100, 30)
    $buttonOpenFile.Add_Click({
        Start-Process "notepad.exe" $outputFile
        $msgBox.Close()
    })
    $msgBox.Controls.Add($buttonOpenFile)

    $buttonOpenFolder = New-Object System.Windows.Forms.Button
    $buttonOpenFolder.Text = "Open Folder"
    $buttonOpenFolder.Location = New-Object System.Drawing.Point(120, 50)
    $buttonOpenFolder.Size = New-Object System.Drawing.Size(100, 30)
    $buttonOpenFolder.Add_Click({
        Start-Process "explorer.exe" "/select,`"$outputFile`""
        $msgBox.Close()
    })
    $msgBox.Controls.Add($buttonOpenFolder)

    $buttonClose = New-Object System.Windows.Forms.Button
    $buttonClose.Text = "Close"
    $buttonClose.Location = New-Object System.Drawing.Point(230, 50)
    $buttonClose.Size = New-Object System.Drawing.Size(50, 30)
    $buttonClose.Add_Click({
        $msgBox.Close()
    })
    $msgBox.Controls.Add($buttonClose)

    # Show dialog
    $msgBox.Add_Shown({$msgBox.Activate()})
    [void] $msgBox.ShowDialog()
}

# Create a folder browser dialog box
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "Select Source Directory"
$folderBrowser.RootFolder = "MyComputer"

# Show the dialog and get the selected folder
$dialogResult = $folderBrowser.ShowDialog()

if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
    $sourceDirectory = $folderBrowser.SelectedPath

    # Check if the entered directory exists
    if (Test-Path -Path $sourceDirectory -PathType Container) {
        # Define the path to the output text file in the selected directory
        $outputFile = Join-Path -Path $sourceDirectory -ChildPath "LibOutput.txt"

        # Ensure the output file exists or create a new one if it doesn't
        if (-not (Test-Path $outputFile)) {
            New-Item -Path $outputFile -ItemType File
        }

        # Clear the output file content before starting to write
        Clear-Content $outputFile

        # Get all music files (.mp3 and .flac) in the directory and subdirectories
        $musicFiles = Get-ChildItem -Path $sourceDirectory -Recurse -Include *.mp3, *.flac

        foreach ($file in $musicFiles) {
            # Use Shell.Application to access the file properties
            $shellObject = New-Object -COMObject Shell.Application
            $folder = $shellObject.NameSpace($file.DirectoryName)
            $fileItem = $folder.ParseName($file.Name)

            # Extract properties: 21 - Title, 13 - Contributing artists
            $title = $folder.GetDetailsOf($fileItem, 21)
            $artist = $folder.GetDetailsOf($fileItem, 13)

           # Write to output file based on metadata availability
           if (-not $title -and -not $artist) {
                # If both title and artist are empty, use the file name
                $file.Name | Out-File -FilePath $outputFile -Append -Encoding UTF8
            } else {
                # If metadata is available, write "Title - Artist"
                if (-not $title) { $title = "Unknown Title" }
                if (-not $artist) { $artist = "Unknown Artist" }
                "$title - $artist" | Out-File -FilePath $outputFile -Append -Encoding UTF8
            }
        }

        # Release the COM object
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shellObject) | Out-Null
        Remove-Variable shellObject
        Show-CompletionDialog -outputFile $outputFile
    } else {
        Write-Host "The selected directory does not exist."
    }
} else {
    Write-Host "No directory selected."
}
