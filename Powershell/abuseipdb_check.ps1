param (
    [string]$inputFile,
    [string]$outputFile,
    [int]$threshold = 20,    # Default threshold for abuse confidence score
    [int]$maxDays = 30,    # Default number of days based on API
    [switch]$help
)
 
# Hardcoded API Key (Replace with your actual API key)
$apiKey = ""   # Replace with your actual AbuseIPDB API key
 
# Show usage information if -help flag is provided
if ($help) {
    Write-Host "Usage: script.ps1 -inputFile <input_file> -outputFile <output_file> [-t <score>] [-d <days>] [-help]"
    Write-Host ""
    Write-Host "Options:"
    Write-Host "  -inputFile <file>   Specify the input CSV file (required)."
    Write-Host "  -outputFile <file>  Specify the output Excel file (required)."
    Write-Host "  -t <score>  	      Specify the minimum abuse confidence score (optional, default is 20)."
    Write-Host "  -md <days>   	      Specify how far back in time we go to fetch reports (optional, 1-365 (30 default))."
    Write-Host "  -help               Show this help message."
    exit
}
 
 
# Ensure both input file and output file are provided
if (-not $inputFile -or -not $outputFile) {
    Write-Host "Error: Both input file and output file are required. Use the -help flag for usage details."
    exit
}
 
# Automatically add .xlsx to the output file if not provided
if (-not $outputFile.EndsWith(".xlsx")) {
	$outputFile = "$outputFile.xlsx"
	Write-Host "Output file converted to Excel format."
}
 
# Ensure the Import-Excel module is installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}
 
# Function to validate IP addresses
function Validate-IP {
    param (
        [string]$IP
    )
    # Regular expression for validating an IP address (IPv4)
    $regex = '^((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$'
    return $IP -match $regex
}
 
# Read the CSV file
$csvData = Import-Csv -Path $inputFile
 
# Extract valid IPs from the CSV
$ipAddresses = $csvData | ForEach-Object {
    $_.PSObject.Properties.Value | ForEach-Object {
        $ip = $_
        if (Validate-IP -IP $ip -and $ip -notmatch '^\d+$') {
            $ip
        }
    }
} | Sort-Object -Unique
 
# Debug: Print extracted IPs
#Write-Host "Extracted IP Addresses:"
#$ipAddresses | ForEach-Object { Write-Host $_ }
 
# Create an array to store results
$results = @()
 
# Process each unique IP address
foreach ($ip in $ipAddresses) {
    #Write-Host "Checking IP: $ip"
 
    try {
    	$apiURL = "https://api.abuseipdb.com/api/v2/check?ipAddress=$ip&verbose=true"
    	if ($maxDays) {
		$apiURL += "&maxAgeInDays=$maxDays"
	}
 
        # Call the AbuseIPDB API
        $response = Invoke-RestMethod -Uri $apiURL `
                                      -Method Get `
                                      -Headers @{ "Key" = $apiKey; "Accept" = "application/json" }
 
        # Debug: Print the API response
        #Write-Host "API Response for ${ip}: " + $($response | ConvertTo-Json -Depth 10)
 
        # Extract details
        $confidenceScore = [int]$response.data.abuseConfidenceScore
	$isp = $response.data.isp
	$usageType = $response.data.usageType
	$country = $response.data.countryName
	$tor  = $response.data.isTor
	$report = $response.data.totalReports
 
        Write-Host "Abuse Confidence Score for ${ip}: Abuse Score: $confidenceScore | Threshold: $threshold | Country: $country"
 
        # Check if the score meets the threshold
        if ($confidenceScore -ge $threshold) {
            $results += [PSCustomObject]@{
                IPAddress      = $ip
                AbuseScore     = $confidenceScore
		ISP	       = $isp
		UsageType      = $usageType
		Location       = $country
		TotalReports   = $report
		Tor_Node       = $tor
                AbuseIPDBLink  = "https://www.abuseipdb.com/check/$ip"
            }
        }
    }
    catch {
        Write-Host "Error processing IP ${ip}: $_"
    }
}
 
# Export results to Excel
if ($results.Count -gt 0) {
    # Delete old file to "overwrite" the data. Comment out if necessary.
    if (Test-Path -Path $outputFile) {
	Remove-Item -Path $outputFile -Force
    } 
    $results | Export-Excel -Path $outputFile -WorksheetName "AbuseIPDB Results" -AutoSize
    Write-Host "Results saved to $outputFile"
} else {
    Write-Host "No IPs with an abuse score above $($threshold) were found."
}