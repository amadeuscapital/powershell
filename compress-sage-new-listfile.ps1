param (
    [string]$rootFolder = "D:\Shares\scratch\SageBackups",
    [Parameter(Mandatory=$true)]
    [string]$targetYearMonth,
    [bool]$test=$false
)

function write_logFile {
    Param ([string]$log_message)
    $logfile = "c:\amit\logs\sage.log"
    $time_stamp = (Get-Date).toString("[yyyy-MM-dd HH:mm:ss]")
    $log_message = "$($time_stamp) $($log_message)"
    [System.IO.File]::AppendAllText($logfile, "$($log_message)`r`n")
}

function cstat ($log) {
    Write-Host -NoNewline ("`r{0,-$([console]::BufferWidth)}" -f $log)
    write_logFile($log)
}

# Get the current year and month
$currentYearMonth = Get-Date -Format "yyyy-MM"

# Check if the targetYearMonth matches the current year and month
if ($targetYearMonth -eq $currentYearMonth) {
    Write-Warning "Specifying the current year and month for archiving is not recommended."
    exit
}

# Get only the files you want (not .7z files)
$filesToArchive = Get-ChildItem -Recurse -File -Path $rootFolder -Filter "*$targetYearMonth*.*" | Where-Object { $_.Extension -ne ".7z" }

# Define the archive name
$archiveName = "$targetYearMonth.7z"
# Define the archive path
$archivePath = Join-Path -Path $rootFolder -ChildPath $archiveName

# Create an array to hold the file paths
$filesToAdd = @()

# Iterate through each file and add it to the list of files to be added
$filesToArchive | ForEach-Object {
    $file = $_
    # Store relative path to rootFolder
    $relativePath = $file.FullName.Substring($rootFolder.Length + 1)
    $filesToAdd += $relativePath
}

# Create a temporary list file
$listFile = Join-Path -Path $rootFolder -ChildPath "tempList.txt"
write-host "List file = $($listFile)"

# Write the relative paths of the files to the list file
$filesToAdd | Out-File -Encoding utf8 -FilePath $listFile

cstat("Archiving files to $($archivePath)")

Push-Location $rootFolder

if (-not $test) {
    # Construct the argument string
    $zipArgs = 'a', "$archiveName", "@$listFile" # Note: Using just $archiveName instead of $archivePath
    write-host $zipArgs
    # Execute the 7-Zip command with the constructed arguments
    & "C:\Program Files\7-Zip\7z.exe" -scsUTF-8 $zipArgs
    
    # Check if the files were added to the archive before deleting
    foreach ($file in $filesToArchive) {
        cstat("Checking $($file.Name) is in $($archiveName)") # Note: Using just $archiveName instead of $archivePath
        $archiveContents = & "C:\Program Files\7-Zip\7z.exe" l "$archiveName"
        if ($archiveContents -match [regex]::Escape($file.Name)) {
            Remove-Item $file.FullName
        }
    }
} else {
    write-host "C:\Program Files\7-Zip\7z.exe a $($archiveName) @$($listFile)"
}

Pop-Location

# Remove the temporary list file
Remove-Item -Path $listFile

