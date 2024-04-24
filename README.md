# Music Metadata Extractor

## Description

This PowerShell script scans specified directories for music files (.mp3 and .flac), extracts metadata (title and artist), and saves this information to a text file (`LibOutput.txt`) in the selected directory. If metadata is missing, the filename is used instead. The script includes a GUI for directory selection and offers post-process options to open the output file, the folder containing it, or to close the dialog box.

## Features

- Extracts music metadata and saves it to a text file.
- Handles .mp3 and .flac file formats.
- Offers GUI-based directory selection.
- Provides options to open the file or folder, or close the application upon completion.

## System Requirements

- Windows 10 or later.
- PowerShell 5.1 or higher.
- .NET Framework 4.5 or higher (for System.Windows.Forms).

## Installation

1. **Download the Script:**
   - Download the `ExtractMusicMetadata.ps1` script from this repository.

2. **Prepare the Batch File:**
   - Create a batch file named `PrintLibToText.bat` in the same directory as the script. This batch file will be used to run the PowerShell script.

## Batch File Setup

Create a new text file and rename it to `PrintLibToText.bat`. Edit it with the following content:

```batch
@echo off
PowerShell -NoProfile -ExecutionPolicy Bypass -File "ExtractMusicMetadata.ps1"
pause
```

Make sure that the path to the PowerShell script is correct. Adjust "ExtractMusicMetadata.ps1" if the script is located in a different folder.

## Running the Script

To run the script:

1. Double-click on PrintLibToText.bat.
2. Follow the on-screen instructions to select the directory and manage the output.

## Contributing

Feel free to fork this project and submit pull requests. You can also open issues if you find bugs or have feature requests.
