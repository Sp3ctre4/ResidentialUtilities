# File Shepherd
# sp3ctre4 | May 5, 2025
# Residential Backup script for gathering all personal files and backing up to an external storage
# Will also scan through the filesystme again and Backup every file into type (Office Docs, Images, PDFs)

Write-Output "====== File Shepherd ====="

### Script Args

# The directory to back up, default is the user's profile
$TargetDirectory = "$env:USERPROFILE"
Write-Output "[o] Targeting Directory: $TargetDirectory"
# where to put the backup
$DestinationDirectory = "D:\Backups"
Write-Output "[o] Destination: $DestinationDirectory"

### Variables
$BackupName= "$(hostname)_Backup_$(Get-Date -Format "yyyy-MM-dd-hh-mm")"
Write-Output "[o] Archive will be stored as: $BackupName"
$LogFile = "$DestinationDirectory\BackupLog.txt"
Write-Output "[o] Script log will be at: $LogFile"
# Office Extensions based on wikipedia: https://en.wikipedia.org/wiki/List_of_Microsoft_Office_filename_extensions
$WordExtensions = ".doc", ".dot", ".wbk", ".docx", ".docm", ".dotx", ".dotm"
$ExcelExtensions = ".xls", ".xlt", ".xlm", ".xlsx", ".xlsm", ".xltx", ".xltm", ".xlsb", ".xla", ".xlam", ".xll", ".xlw", ".xll_", "xla_", ".xla5", ".xla8"
$PowerPointExtensions = ".ppt", ".pot", ".pps", ".ppa", ".pptx", "pptm", ".potx", ".potm", ".ppam", ".ppsx", ".ppsm", ".sldx", ".sldm", ".ppam"
$MiscExtensions = ".accda", ".accdb", ".accde", ".accdr", ".accdt", ".accdu", ".one", ".ecf", ".pub", ".pdf", ".wll", ".wwl", ".csv"
# Most common image extensions
$RasterImageExtensions = ".jpg", ".jpeg", ".jpe", ".png", ".gif", ".webp", ".tiff", ".psd", ".raw", ".bmp", ".hief", ".hiec", ".indd", ".jp2", ".svg"
$VectorImageExtensions = ".svg", ".svgz", ".ai", ".eps"

### Code

if( ! (Test-Path -Path "$DestinationDirectory")){
    New-Item -Path "$DestinationDirectory" -ItemType Directory | Out-Null
    Write-Output "[-] Created output directory: $DestinationDirectory"
}

# Clear Error stream
$Error.Clear()

# create a compressed archive of the entire directory
Write-Output "[-] Compressing Target Directory: $TargetDirectory..."
Compress-Archive -Path $TargetDirectory -DestinationPath "$DestinationDirectory\$BackupName.zip" -CompressionLevel Fastest -Verbose -ErrorAction SilentlyContinue

# Will search the TargetDirectory for any of the specified files (by extension array) and copy them into the tdestination dir (targetdir)
Write-Output "[-] Beginning File Organization...."
function Backup-Files {
    param([string]$Message, [array]$ExtensionList, [string]$TargetDir)
    if ( ! (Test-Path -Path "$TargetDir")) {
        New-Item -Path "$TargetDir" -ItemType Directory | Out-Null
        Write-Output "[-] Created backup directory: $TargetDir"
    }
    Write-Output $Message
    foreach($extension in $ExtensionList){
        Get-ChildItem -Path $TargetDirectory -Filter "*$extension" -Recurse | Copy-Item -Destination $TargetDir -Verbose
    }
}

# Search for Office Files
Backup-Files -Message "[-] Now searching for Word files..." -ExtensionList $WordExtensions -TargetDir "$DestinationDirectory\WordDocuments"
Backup-Files -Message "[-] Now searching for Excel files..." -ExtensionList $ExcelExtensions -TargetDir "$DestinationDirectory\ExcelDocuments"
Backup-Files -Message "[-] Now searching for PowerPoint files..." -ExtensionList $PowerPointExtensions -TargetDir "$DestinationDirectory\PowerPoints"
Backup-Files -Message "[-] Now searching for PDFs, OneNotes, Publisher Docs, and other misc Office files..." -ExtensionList $MiscExtensions -TargetDir "$DestinationDirectory\MiscOfficeDocuments"
# Search for Images
Backup-Files -Message "[-] Now searching for Raster Images..." -ExtensionList $RasterImageExtensions -TargetDir "$DestinationDirectory\RasterImages"
Backup-Files -Message "[-] Now searching for Vector Images..." -ExtensionList $VectorImageExtensions -TargetDir "$DestinationDirectory\VectorImages"

# Handle Errors (if any, pipe to log file)
if ($Error) {
    Write-Output $Error > "$LogFile"
    Write-Output "[!] Errors detected, see log file at: $LogFile"
}
else {
    Write-Output "[-] No errors detected."
}

# Exit the script
Write-Output "===== Script Complete ====="
