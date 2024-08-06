# Function to get the page count from a Word document
function Get-WordDocPageCount {
    param (
        [string]$filePath, # Path to the Word document
        [ref]$wordApp        # Reference to the Word application COM object
    )

    try {
        # Check if the Word application object is initialized
        if ($null -eq $wordApp.Value) {
            Write-Host "Error: Word application object is not initialized."
            return 0
        }

        # Open the Word document
        $doc = $wordApp.Value.Documents.Open($filePath)
        if ($null -eq $doc) {
            Write-Host "Error: Unable to open Word document at path: $filePath"
            return 0
        }

        # Get the page count
        $pageCount = $doc.ComputeStatistics([Microsoft.Office.Interop.Word.WdStatistic]::wdStatisticPages)
        $doc.Close()  # Close the document
        return $pageCount
    }
    catch {
        # Handle exceptions
        Write-Host "Error processing Word document at path: $filePath"
        Write-Host "Exception: $_"
        return 0
    }
}

# Function to get the page count from a PDF document
function Get-PDFPageCount {
    param (
        [string]$filePath  # Path to the PDF document
    )

    # Validate file path
    if ([string]::IsNullOrEmpty($filePath)) {
        Write-Host "Error: File path is empty or null."
        return 0
    }

    try {
        # Read the content of the PDF file
        $content = Get-Content -Path $filePath -Raw -ReadCount 0
        # Count the number of pages based on regex matches
        $matches = [regex]::Matches($content, "/Type\s*/Page[^s]")
        return $matches.Count
    }
    catch {
        # Handle exceptions
        Write-Host "Error processing PDF document at path: $filePath"
        Write-Host "Exception: $_"
        return 0
    }
}

# Function to add data to a CSV file
function Add-DataToCSV {
    param (
        [array]$dataArray, # Array of objects containing file data
        [string]$csvPath     # Path to the CSV file
    )

    foreach ($data in $dataArray) {
        # Format data as CSV line and append to file
        $csvLine = "$($data.FileName),$($data.Pages),$($data.Type)"
        $csvLine | Out-File -Append $csvPath
    }
}

# Function to process folder and update CSV with document page counts
function Get-FolderPageCounts {
    param (
        [string]$folderPath, # Path to the folder containing documents
        [bool]$writeCSV = $false, # Flag to toggle CSV writing
        [bool]$includeDateInFileName = $false,  # Flag to include the current date in the CSV file name
        [bool]$includeSummary = $false  # Flag to display summary information
    )

    # Prompt for folder path if not provided
    if (-not $folderPath) {
        $folderPath = Read-Host "Drag a folder or leave blank for the default folder"
        if (-not $folderPath) {
            $folderPath = "./Default"  # Default folder path
        }
    }

    $dataArray = @()  # Array to store document data

    # Get lists of Word and PDF files
    $wordFiles = Get-ChildItem -Path $folderPath -Filter "*.docx" -Recurse -ErrorAction SilentlyContinue
    $pdfFiles = Get-ChildItem -Path $folderPath -Filter "*.pdf" -Recurse -ErrorAction SilentlyContinue

    try {
        # Create Word application COM object
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
    }
    catch {
        # Handle exceptions
        Write-Host "Error: Failed to create Word application object."
        Write-Host "Exception: $_"
        return
    }

    Write-Host "`n--------------------------------------------------------------------------------------------------------"

    # Process Word files
    foreach ($wordFile in $wordFiles) {
        $wordPageCount = Get-WordDocPageCount -filePath $wordFile.FullName -wordApp ([ref]$word)
        $dataArray += [pscustomobject]@{
            FileName = $wordFile.Name
            Pages    = $wordPageCount
            Type     = "Word Document"
        }
        Write-Host "File: $($wordFile.Name) - Pages: $wordPageCount - Type: Word Document"
    }
    $word.Quit()

    # Process PDF files
    foreach ($pdfFile in $pdfFiles) {
        $pdfPageCount = Get-PDFPageCount -filePath $pdfFile.FullName
        $dataArray += [pscustomobject]@{
            FileName = $pdfFile.Name
            Pages    = $pdfPageCount
            Type     = "PDF Document"
        }
        Write-Host "File: $($pdfFile.Name) - Pages: $pdfPageCount - Type: PDF Document"
    }

    # Display summary information if flag is set    
    if ($includeSummary) {
        Write-Host "`n-------------------------------------------------Summary------------------------------------------------"
        # Calculate and display the total number of Word pages
        $totalWordPages = ($dataArray | Where-Object { $_.Type -eq 'Word Document' } | Measure-Object -Property Pages -Sum).Sum
        Write-Host "Total Word Pages: $totalWordPages"

        # Calculate and display the total number of PDF pages
        $totalPdfPages = ($dataArray | Where-Object { $_.Type -eq 'PDF Document' } | Measure-Object -Property Pages -Sum).Sum
        Write-Host "Total PDF Pages: $totalPdfPages"
        
        Write-Host "Total Word files processed: $($wordFiles.Count)"
        Write-Host "Total PDF files processed: $($pdfFiles.Count)"
        Write-Host "Total files processed: $($wordFiles.Count + $pdfFiles.Count)"
        Write-Host "--------------------------------------------------------------------------------------------------------"
    }

    # Handle CSV file writing based on flag
    if ($writeCSV) {
        # Determine the CSV file path

        # Generate CSV file path with current date in YYYYMMDD format
        if ($includeDateInFileName) {
            $csvPath = "$env:USERPROFILE\Downloads\document_page_counts_" + (Get-Date -Format "yyyyMMdd") + ".csv"
        }
        else {
            $csvPath = "$env:USERPROFILE\Downloads\document_page_counts.csv"
        }
        # Handle existing CSV file scenarios
        if (Test-Path $csvPath) {
            $choice = Read-Host "`nThe CSV file already exists. Choose an option:`n1. Append`n2. Overwrite`n3. Cancel`n`n(Enter 1, 2, or 3)"

            switch ($choice) {
                "1" {
                    # Continue with existing file, no action needed
                }
                "2" {
                    # Overwrite file and add header
                    $csvData = "File,Pages,Type"
                    $csvData | Out-File $csvPath
                }
                "3" {
                    Write-Host "Operation cancelled."
                    return
                }
                default {
                    Write-Host "Invalid choice. Operation cancelled."
                    return
                }
            }
        }
        else {
            # Create new file and add header
            $csvData = "File,Pages,Type"
            $csvData | Out-File $csvPath
        }

        # Add collected data to the CSV file
        Add-DataToCSV -dataArray $dataArray -csvPath $csvPath
        Write-Host "CSV file updated at: $csvPath"
    }
    else {
        Write-Host "CSV writing is disabled. No CSV file was created or updated."
    }
}

# Run the function to process folder and optionally update CSV
# Parameters: folderPath, writeCSV, includeDateInFileName, summary
# folderPath: Path to the folder containing documents (optional) - omit to use default folder
# writeCSV: Flag to toggle CSV writing (optional) - default is true
# includeDateInFileName: Flag to include the current date in the CSV file name (optional) - default is false
# includeSummary: Flag to display summary information (optional) - default is false
# Example: Get-FolderPageCounts -writeCSV $true -folderPath "C:\Documents" -includeDateInFileName $true -includeSummary $true

Get-FolderPageCounts -writeCSV $true -folderPath "" -includeDateInFileName $true -includeSummary $true
