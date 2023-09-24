<#
    .SYNOPSIS
        Merges all csv files in a folder into one csv file.

    .PARAMETER Path
        The path to the folder containing the csv files.

    .PARAMETER OutputFile
        The path to the output file.
        If not provided, the output file will be created in the same folder as the input files.

    .EXAMPLE
        Merge-CsvFiles -Path "C:\Temp\csv-files" -OutputFile "C:\Temp\merged.csv"

        This example merges all csv files in the folder "C:\Temp\csv-files" into the file "C:\Temp\merged.csv".

    .EXAMPLE
        Merge-CsvFiles -Path "C:\Temp\csv-files"

        This example merges all csv files in the folder "C:\Temp\csv-files" into a file named "merged-<timestamp>.csv" in the same folder.
#>

# Merge Csv Files
[CmdletBinding()]
param(
    [string]$Path = $PSScriptRoot,

    [ValidateScript({$_ -like "*.csv"})]
    [string]$OutputFile = (Join-Path -Path $Path -ChildPath "merged-$(Get-Date -Format yyyyMMddhhmmss).csv")
)

# Check if path exists
if(-not (Test-Path -Path $Path)) {
    Write-Error "Path $Path does not exist"
    return
}

# Get all csv files
$files = Get-ChildItem -Path $Path -Filter *.csv

# Check if files exist
if($files.Count -eq 0) {
    Write-Error "No csv files found in $Path"
    return
}

# Prepare output file path, if not provided
if($OutputFile -notlike "*\*") {
    $OutputFile = Join-Path -Path $Path -ChildPath $OutputFile
}

# Check if output file exists
if(Test-Path -Path $OutputFile) {
    Write-Error "Output file $OutputFile already exists"
    return
}

# Create output file
$null = New-Item -Path $OutputFile -ItemType File -Force | Out-Null

# Export header
$null = Get-Content -Path $files[0].FullName | Select-Object -First 1 | Add-Content -Path $OutputFile

# Export content
$progressCounter = 0
foreach($file in $files) {
    $file = $file.FullName
    Write-Progress -Activity "Merging csv files" -Status "Processing $file" -PercentComplete (($progressCounter / $files.Count) * 100)
    $null = Get-Content -Path $file | Select-Object -Skip 1 | Add-Content -Path $OutputFile
    Start-Sleep -Seconds 1
    $progressCounter++
}