# Define the base folder and script paths at the beginning of the script
$baseFolderPath = "C:\Power BI Backups"

# Set the workspace ID
$dataflowworkspaceId = "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX"

# Define the report backups path
$dataflowBackupsPath = Join-Path -Path $baseFolderPath -ChildPath "Dataflow Backups"

# Check and set the execution policy
$currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
if ($currentPolicy -eq 'Restricted' -or $currentPolicy -eq 'Undefined' -or $currentPolicy -eq 'AllSigned') {
    Write-Host "Current execution policy is restrictive: $currentPolicy. Attempting to set to RemoteSigned."
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
    Write-Host "Execution policy set to RemoteSigned."
} else {
    Write-Host "Current execution policy is sufficient: $currentPolicy."
}

# Check and install the necessary modules
$requiredModules = @('MicrosoftPowerBIMgmt', 'ImportExcel')
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Install-Module -Name $module -Scope CurrentUser -Force
    }
}

# Check if the base folder exists, if not create it
if (-not (Test-Path -Path $baseFolderPath)) {
    New-Item -Path $baseFolderPath -ItemType Directory
}

# Check if the "dataflow Backups" folder exists, if not create it
if (-not (Test-Path -Path $dataflowBackupsPath)) {
    New-Item -Path $dataflowBackupsPath -ItemType Directory
}

# Create a variable for end of week (Friday) date
$date = (Get-Date -UFormat "%Y-%m-%d")

# Create a new folder for the backups
$dataflow_new_date_folder = Join-Path -Path $dataflowBackupsPath -ChildPath $date
New-Item -Path $dataflow_new_date_folder -ItemType Directory -Force

# Import the required module
Import-Module MicrosoftPowerBIMgmt

# Authenticate to Power BI Service
Connect-PowerBIServiceAccount

# Set the base output file path
$baseOutputFilePath = $dataflow_new_date_folder

# Set the combined Excel output path
$combinedExcelOutputPath = Join-Path -Path $dataflow_new_date_folder -ChildPath "DataflowDetail.xlsx"

# Set the Power BI REST API URL for the dataflow details
$dataflowDetailsUrl = "https://api.powerbi.com/v1.0/myorg/groups/$dataflowworkspaceId/dataflows"

# Get the list of dataflows in the workspace
$dataflowsResponse = Invoke-PowerBIRestMethod -Url $dataflowDetailsUrl -Method Get

# Parse the JSON response
$dataflows = $dataflowsResponse | ConvertFrom-Json

# Initialize a combined DataTable
$combinedDataTable = New-Object System.Data.DataTable
$combinedDataTable.Columns.Add("Dataflow ID", [System.String])
$combinedDataTable.Columns.Add("Dataflow Name", [System.String])
$combinedDataTable.Columns.Add("Query Name", [System.String])
$combinedDataTable.Columns.Add("Query", [System.String])
$combinedDataTable.Columns.Add("Report Date", [System.DateTime])

# Get the current date
$currentDate = Get-Date

# Function to check if a position is within curly braces
function IsInsideCurlyBraces {
    param (
        [string]$text,
        [int]$position
    )
    $openBraces = 0
    for ($i = 0; $i -lt $position; $i++) {
        if ($text[$i] -eq '{') { $openBraces++ }
        elseif ($text[$i] -eq '}') { $openBraces-- }
    }
    return $openBraces -gt 0
}

# Check if the response is valid and contains dataflows
if ($dataflows -and $dataflows.value) {
    Write-Host "Dataflows found: $($dataflows.value.Count)"
    
    # Iterate through the dataflows
    foreach ($dataflow in $dataflows.value) {
        $dataflowId = $dataflow.objectId
        $dataflowName = $dataflow.name
        
        # Define output file path specific to the dataflow
        $dataflowOutputFilePath = Join-Path -Path $baseOutputFilePath -ChildPath "$dataflowName-dataflow.txt"
        
        # Set the Power BI REST API URL for the specific dataflow
        $apiUrl = "https://api.powerbi.com/v1.0/myorg/groups/$dataflowworkspaceId/dataflows/$dataflowId"
        
        # Get the dataflow
        $response = Invoke-PowerBIRestMethod -Url $apiUrl -Method Get
        
        # Convert the response to JSON string
        $jsonString = $response | ConvertTo-Json
        
        # Write the JSON string to a text file
        $jsonString | Out-File -FilePath $dataflowOutputFilePath -Encoding UTF8
        
        # Extract the data from the JSON response without writing intermediate files
        $startMarker = '"document\":'
        $endMarker = '"connectionOverrides\":'
        $startIndex = $jsonString.IndexOf($startMarker) + $startMarker.Length
        $endIndex = $jsonString.IndexOf($endMarker, $startIndex)
        $documentContent = $jsonString.Substring($startIndex, $endIndex - $startIndex)

        # Format the extracted content
        $formattedText = $documentContent -replace '\\r\\n', "`n" `
                                                -replace '\\\"', '"' `
                                                -replace '\\\\', '\' `
                                                -replace '(?<=\w)(=|then|else)(?=\w)', ' $1 ' `
                                                -replace '(?<=then)\s+', ' ' `
                                                -replace '\s+(?=else)', "`n    "

        # Additional formatting
        $insideQuotes = $false
        $formattedTextStep7 = ""
        for ($i = 0; $i -lt $formattedText.Length; $i++) {
            $char = $formattedText[$i]
            if ($char -eq '"') {
                $insideQuotes = -not $insideQuotes
            }
            if ($char -eq ',' -and -not $insideQuotes -and -not (IsInsideCurlyBraces -text $formattedText -position $i)) {
                $formattedTextStep7 += "$char`n    "
            } else {
                $formattedTextStep7 += $char
            }
        }

        $formattedTextStep8 = $formattedTextStep7 -replace 'nshared', 'QueryStartandEndMarker' `
                                                  -replace '\\r', ' ' `
                                                  -replace '\\n', ' ' `
                                                  -replace '\\', '' `
                                                  -replace '\r\n', "`n" `
                                                  -creplace '(?<!["\w])r(\s)(?![\w"])', '$1' `
                                                  -replace '\;r', ''

        $formattedTextStep9 = $formattedTextStep8 -replace '(?<=let|in|each)\s+', "`n    " `
                                                  -replace '(?<=,\s*)#', "`n    #"

        $formattedTextStep9 = $formattedTextStep9 -replace '(?<!\n)QueryStartandEndMarker\s', "`nQueryStartandEndMarker`n`n"

        $formattedTextStep9 = $formattedTextStep9 -replace '(^|\s)let', "`nlet"

	$formattedTextStep9 = $formattedTextStep9 -replace '\)  in', ")`n    in"

        # Read the content directly from the formatted text
        $fileContent = $formattedTextStep9

        # Initialize variables
        $inQuery = $false
        $queryName = ""
        $query = ""
        $data = @()

        # Split the content into lines
        $lines = $fileContent -split "`n"

        # Iterate over the lines
        foreach ($line in $lines) {
            if ($line -match 'QueryStartandEndMarker\s*') {
                # If we find a new query, save the previous one
                if ($queryName -ne "") {
                    $data += [PSCustomObject]@{
                        "Dataflow ID" = $dataflowId
                        "Dataflow Name" = $dataflowName
                        "Query Name" = $queryName
                        "Query" = $query.Trim()
                        "Report Date" = $currentDate
                    }
                }
                # Reset variables for new query
                $inQuery = $true
                $queryName = ""
                $query = ""
            } elseif ($inQuery -and $line.Trim() -ne "") {
                # Set the query name as the first non-empty line after QueryStartandEndMarker
                if ($queryName -eq "") {
                    $queryName = $line.Trim()
                } else {
                    # Append the line to the query
                    $query += "$line`n"
                }
            } elseif ($queryName -ne "") {
                # Append the line to the query
                $query += "$line`n"
            }
        }

        # Add the last query
        if ($queryName -ne "") {
            $data += [PSCustomObject]@{
                "Dataflow ID" = $dataflowId
                "Dataflow Name" = $dataflowName
                "Query Name" = $queryName
                "Query" = $query.Trim()
                "Report Date" = $currentDate
            }
        }

        # Fill the combined DataTable with data
        foreach ($item in $data) {
            $row = $combinedDataTable.NewRow()
            $row["Dataflow ID"] = $item."Dataflow ID"
            $row["Dataflow Name"] = $item."Dataflow Name"
            $row["Query Name"] = $item."Query Name"
            $row["Query"] = $item.Query
            $row["Report Date"] = $item."Report Date"
            $combinedDataTable.Rows.Add($row)
        }
    }

    # Export the combined DataTable to an Excel file
    $combinedDataTable | Export-Excel -Path $combinedExcelOutputPath -AutoSize
    Write-Output "Data exported to $combinedExcelOutputPath"

} else {
    Write-Host "No dataflows found or an error occurred."
    Write-Host "Response: $($dataflowsResponse | ConvertTo-Json -Depth 10)"
}

# Combine files if both there
$fileName = "DataflowDetail.xlsx"
$sourceFilePath = Join-Path -Path $dataflow_new_date_folder -ChildPath $fileName
$destinationFilePath = Join-Path -Path $baseFolderPath -ChildPath $fileName

# Check if the source file exists
if (-not (Test-Path -Path $sourceFilePath)) {
    Write-Error "Source file not found: $sourceFilePath"
    exit
}

# Check if the destination file exists
if (Test-Path -Path $destinationFilePath) {
    # Load the source and destination Excel files
    $sourceData = Import-Excel -Path $sourceFilePath
    $destinationData = Import-Excel -Path $destinationFilePath
    
    # Combine the data
    $combinedData = $destinationData + $sourceData
    
    # Export the combined data to the destination file
    $combinedData | Export-Excel -Path $destinationFilePath -WorksheetName "Sheet1"
} else {
    # If the file doesn't exist in the destination folder, copy it there
    Copy-Item -Path $sourceFilePath -Destination $destinationFilePath
}

Write-Output "Excel files processed and combined successfully."
