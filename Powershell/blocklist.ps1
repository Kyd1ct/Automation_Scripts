param (
    [string]$inputFilePath,
    [string]$outputFilePath,
    [switch]$excel,
    [switch]$rundeck,
    [switch]$help
)
 
# Show usage information if -h flag is provided
if ($help) {
    Write-Host "Usage: script.ps1 -i <input_file> [-o <output_file>] [-e] [-r] [-h]"
    Write-Host ""
    Write-Host "Options:"
    Write-Host "  -i <input_file>   Specify the input file (required)."
    Write-Host "  -o <output_file>  Specify the output file (optional). Will be displayed in the console if not provided"
    Write-Host "  -e                Specify file type as Excel"
    Write-Host "  -r                Display unique IPs in a single line separated by ';' for direct Rundeck input."
    Write-Host "  -h                Show this help message."
    exit
}
 
# Input file provision checker
if (-not $inputFilePath) {
    Write-Host "Error: Input file is required. Use the -h flag for help."
    exit
}
 
# Define the URL for blocklist | Change where necessary.
$blocklistUrl = "URL"
 
# Initialize $inputIPs
$inputIPs = @()
 
if ($excel) {
    # Process Excel file input
    try {
        # Ensure the ImportExcel module is available
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            Write-Host "Error: ImportExcel module is required to process Excel files. Install it using: Install-Module -Name ImportExcel"
            exit
        }
 
        # Import the Excel module
        Import-Module ImportExcel
 
        # Read first column, skipping the first row, and omit non-IP strings
        $inputIPs = Import-Excel -Path $inputFilePath | ForEach-Object { $_.IPAddress } `
            | Select-Object -Skip 1 `
            | ForEach-Object { ($_ -replace '\s+', '').Trim() } `
            | Where-Object { $_ -match '\b(?:\d{1,3}\.){3}\d{1,3}\b' }
 
    } catch {
        Write-Host "Error: Failed to process the Excel file. $_"
        exit
    }
} else {
    # Process text file input
    $inputIPs = Get-Content -Path $inputFilePath | ForEach-Object { ($_ -replace '\s+', '').Trim() }
}
 
# Fetch the content from blocklist URL using Invoke-WebRequest
$response = Invoke-WebRequest -Uri $blocklistUrl
$blocklistIPs = $response.Content -split "`n" | ForEach-Object { ($_ -replace '\s+', '').Trim() }
 
# Ensure both arrays are treated as arrays of strings to ensure that diff will work
$inputIPs = $inputIPs | ForEach-Object { [string]$_ }
$blocklistIPs = $blocklistIPs | ForEach-Object { [string]$_ }
 
# Compare the IPs and filter out the ones that are not in blocklist
$uniqueIPs = $inputIPs | Where-Object { $_ -notin $blocklistIPs }
 
# Generate Rundeck output based on the -r switch
$outputContent = if ($rundeck) {
    $uniqueIPs -join ";"
} else {
    # Default output: Line by line
    $uniqueIPs
}
 
# If output file path is provided, write the result to the file
if ($outputFilePath) {
    # Check if the output file path includes a directory, if not, use the current working directory
    if (-not [System.IO.Path]::IsPathRooted($outputFilePath)) {
        $outputFilePath = Join-Path -Path (Get-Location) -ChildPath $outputFilePath
    }
 
    # Check if the output directory exists, if not, create it
    $outputDir = [System.IO.Path]::GetDirectoryName($outputFilePath)
    if (-not (Test-Path -Path $outputDir)) {
        New-Item -ItemType Directory -Force -Path $outputDir
    }
 
    # Write the output content to the file | "-Force" WILL overwrite the file!
    $outputContent | Out-File -FilePath $outputFilePath -Force
} else {
    # Display the result in the console
    if ($rundeck) {
        Write-Host "Unique IPs (Rundeck-style):"
        Write-Host $outputContent
    } else {
        Write-Host "Unique IPs (not found in blocklist):"
        $outputContent | ForEach-Object { Write-Host $_ }
    }
}