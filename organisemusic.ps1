# Function to get ID3 tags using Shell.Application COM object
function Get-MP3ID3Tags {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    # Create a COM object for Shell
    $shell = New-Object -ComObject Shell.Application

    # Extract the folder and file info using -LiteralPath
    $folder = $shell.Namespace((Get-Item -LiteralPath $FilePath).DirectoryName)
    $file = $folder.ParseName((Get-Item -LiteralPath $FilePath).Name)

    # Get the metadata properties
    $title = $folder.GetDetailsOf($file, 21)  # Title field (index 21)
    $album = $folder.GetDetailsOf($file, 14)  # Album field (index 14)
    $artists = $folder.GetDetailsOf($file, 13) # Contributing artists (index 13)

    return @{
        Title  = $title
        Album  = $album
        Artists = $artists
        FilePath = $FilePath
    }
}

# Open a folder dialog to choose the source folder containing MP3 files
Add-Type -AssemblyName System.Windows.Forms

# Initialize the FolderBrowserDialog for source folder
$FolderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
$FolderBrowserDialog.Description = "Select the folder containing MP3 files"
$FolderBrowserDialog.ShowNewFolderButton = $true

# Show the dialog and get the selected folder path for the source
Write-Host "Step 1: Please select the folder containing MP3 files..."
$dialogResult = $FolderBrowserDialog.ShowDialog()

# Check if a source folder was selected
if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
    $sourceFolderPath = $FolderBrowserDialog.SelectedPath
    Write-Host "Source folder selected: $sourceFolderPath"

    # Open another dialog to select the destination folder
    $FolderBrowserDialog.Description = "Select the destination folder"
    Write-Host "Step 2: Please select the destination folder..."
    $dialogResult = $FolderBrowserDialog.ShowDialog()

    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        $destinationFolderPath = $FolderBrowserDialog.SelectedPath
        Write-Host "Destination folder selected: $destinationFolderPath"

        # Get all MP3 files in the selected source folder
        $mp3Files = Get-ChildItem -Path $sourceFolderPath -Filter *.mp3

        # Check if any MP3 files were found
        if ($mp3Files.Count -eq 0) {
            Write-Host "No MP3 files found in the selected source folder."
        } else {
            Write-Host "Processing ${mp3Files.Count} MP3 files..."
            # Loop through each MP3 file and extract metadata
            foreach ($mp3File in $mp3Files) {
                $metadata = Get-MP3ID3Tags -FilePath $mp3File.FullName
                
                # Create destination folders based on contributing artists and album
                $artistFolder = Join-Path -Path $destinationFolderPath -ChildPath $metadata.Artists
                $albumFolder = Join-Path -Path $artistFolder -ChildPath $metadata.Album
                
                # Create the folders if they do not exist
                if (-not (Test-Path -LiteralPath $artistFolder)) {
                    New-Item -ItemType Directory -Path $artistFolder | Out-Null
                    Write-Host "Created artist folder: $artistFolder"
                }
                if (-not (Test-Path -LiteralPath $albumFolder)) {
                    New-Item -ItemType Directory -Path $albumFolder | Out-Null
                    Write-Host "Created album folder: $albumFolder"
                }

                # Move the MP3 file to the new location
                $destinationFilePath = Join-Path -Path $albumFolder -ChildPath $mp3File.Name
                Move-Item -Path $mp3File.FullName -Destination $destinationFilePath -Force

                # Output the action taken
                Write-Host "Moved '$($metadata.Title)' to '$destinationFilePath'"
            }
            Write-Host "All MP3 files processed successfully."
        }
    } else {
        Write-Host "Destination folder selection was canceled. Exiting script."
    }
} else {
    Write-Host "Source folder selection was canceled. Exiting script."
}
