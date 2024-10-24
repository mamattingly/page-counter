# Function to get the page count from a Word document
function Get-WordDocPageCount {
    param (
        [string]$filePath, # Path to the Word document
        [ref]$wordApp      # Reference to the Word application COM object
    )

    try {
        # Check if the Word application object is initialized
        if (-not $wordApp.Value) {
            Write-Host "Error: Word application object is not initialized."
            return 0
        }

        # Open the Word document
        $doc = $wordApp.Value.Documents.Open($filePath)
        if (-not $doc) {
            Write-Host "Error: Unable to open Word document at path: $filePath"
            return 0
        }

        # Get the page count and close the document
        $pageCount = $doc.ComputeStatistics([Microsoft.Office.Interop.Word.WdStatistic]::wdStatisticPages)
        $doc.Close()
        return $pageCount
    }
    catch {
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
        # Read the content of the PDF file and count the pages
        $content = Get-Content -Path $filePath -Raw -ReadCount 0
        $matches = [regex]::Matches($content, "/Type\s*/Page[^s]")
        return $matches.Count
    }
    catch {
        Write-Host "Error processing PDF document at path: $filePath"
        Write-Host "Exception: $_"
        return 0
    }
}

# Function to add data to a CSV file
function Add-DataToCSV {
    param (
        [array]$dataArray,  # Array of objects containing file data
        [string]$csvPath     # Path to the CSV file
    )

    foreach ($data in $dataArray) {
        # Format data as CSV line and append to file
        "$($data.FileName),$($data.Pages),$($data.Type)" | Out-File -Append $csvPath
    }
}

# Function to process folder and update CSV with document page counts
function Get-FolderPageCounts {
    param (
        [string]$folderPath = "$PSScriptRoot/Page Counter", # Default folder path
        [bool]$writeCSV = $false,                             # Flag to toggle CSV writing
        [bool]$includeDateInFileName = $false,               # Flag to include date in CSV file name
        [bool]$includeSummary = $false,                       # Flag to display summary information
        [bool]$silentMode = $true,                            # Flag to suppress prompts
        [string]$toEmailAddress                                # Recipient email address
    )

    # Prompt for folder path if not provided and not in silent mode
    if (-not $folderPath -and -not $silentMode) {
        $folderPath = Read-Host "Drag a folder or leave blank for the default folder"
    }

    $dataArray = @()  # Array to store document data

    # Get lists of Word and PDF files
    $wordFiles = Get-ChildItem -Path $folderPath -Filter "*.docx" -Recurse -ErrorAction SilentlyContinue
    $pdfFiles = Get-ChildItem -Path $folderPath -Filter "*.pdf" -Recurse -ErrorAction SilentlyContinue

    # Create Word application COM object
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
    }
    catch {
        Write-Host "Error: Failed to create Word application object."
        Write-Host "Exception: $_"
        return
    }

    if (-not $silentMode) {
        Write-Host "`n--------------------------------------------------------------------------------------------------------"
    }

    # Process Word files
    foreach ($wordFile in $wordFiles) {
        $wordPageCount = Get-WordDocPageCount -filePath $wordFile.FullName -wordApp ([ref]$word)
        $dataArray += [pscustomobject]@{
            FileName = $wordFile.Name
            Pages    = $wordPageCount
            Type     = "Word Document"
        }
        if (-not $silentMode) {
            Write-Host "File: $($wordFile.Name) - Pages: $wordPageCount - Type: Word Document"
        }
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
        if (-not $silentMode) {
            Write-Host "File: $($pdfFile.Name) - Pages: $pdfPageCount - Type: PDF Document"
        }
    }

    # Display summary information if flag is set    
    $totalWordPages = ($dataArray | Where-Object { $_.Type -eq 'Word Document' } | Measure-Object -Property Pages -Sum).Sum
    $totalPdfPages = ($dataArray | Where-Object { $_.Type -eq 'PDF Document' } | Measure-Object -Property Pages -Sum).Sum
    $totalPages = $totalWordPages + $totalPdfPages

    if ($includeSummary -and -not $silentMode) {
        Write-Host "`n-------------------------------------------------Summary------------------------------------------------"
        Write-Host "Total Word Pages: $totalWordPages"
        Write-Host "Total PDF Pages: $totalPdfPages"
        Write-Host "Total Word files processed: $($wordFiles.Count)"
        Write-Host "Total PDF files processed: $($pdfFiles.Count)"
        Write-Host "Total files processed: $($wordFiles.Count + $pdfFiles.Count)"
        Write-Host "--------------------------------------------------------------------------------------------------------"
    }

    # Handle CSV file writing based on flag
    if ($writeCSV) {
        $csvPath = "$PSScriptRoot/"

        # Generate CSV file path with current date if needed
        $csvPathFull = if ($includeDateInFileName) {
            $csvPath + "document_page_counts_" + (Get-Date -Format "yyyyMMdd") + ".csv"
        } else {
            $csvPath + "document_page_counts.csv"
        }

        # Handle existing CSV file scenarios
        if (Test-Path $csvPathFull) {
            if (-not $silentMode) {
                $choice = Read-Host "`nThe CSV file already exists. Choose an option:`n1. Append`n2. Overwrite`n3. Cancel`n`n(Enter 1, 2, or 3)"
                switch ($choice) {
                    "1" { }  # Continue with existing file
                    "2" {
                        # Overwrite file and add header
                        "File,Pages,Type" | Out-File $csvPathFull
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
                # Overwrite file and add header
                "File,Pages,Type" | Out-File $csvPathFull
            }
        }

        # Add collected data to the CSV file
        Add-DataToCSV -dataArray $dataArray -csvPath $csvPathFull
        if (-not $silentMode) {
            Write-Host "`nCSV file updated at: $csvPathFull"
        }
    }
    else {
        if (-not $silentMode) {
            Write-Host "`nCSV writing is disabled. No CSV file was created or updated."
        }
    }
    
    Write-Host "Operation Completed Successfully"
    Send-Email -totalPages $totalPages -toEmailAddress $toEmailAddress -dataArray $dataArray
}

# Function to send email with attachment
function Send-Email {
    param (
        [int]$totalPages,  # Total number of pages processed
        [string]$toEmailAddress,
        [array]$dataArray  # Array of objects containing file data
    )

    # Create an Outlook COM object
    $currentDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)

    # Construct the email body
    $emailBody = "Page Counter Results - $currentDateTime`n"
    $emailBody += "Total Pages: $totalPages`n`n"
    $emailBody += "File, Pages, Type`n--------------------------------------`n"

    foreach ($data in $dataArray) {
        $emailBody += "Name: $($data.FileName), Page Count: $($data.Pages), Document Type: $($data.Type)`n"
    }

    # Set the email properties
    $mail.Subject = "Page Counter Results - $currentDateTime"
    $mail.Body = $emailBody
    $mail.To = $toEmailAddress

    # Send the email
    $mail.Send()
    Write-Host "Email sent to: $toEmailAddress"
}

Get-FolderPageCounts -writeCSV $false -folderPath "" -includeDateInFileName $false -includeSummary $false -toEmailAddress "mikeamatt@hotmail.com"
