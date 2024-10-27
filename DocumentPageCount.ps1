Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$configFilePath = "$PSScriptRoot/config.json"

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

# Function to process folder and update CSV with document page counts
function Get-FolderPageCounts {
    param (
        [string]$folderPath, # Default folder path
        [string]$toEmailAddress, # Email address to send reports to
        [bool]$writeCSV, # Flag to toggle CSV writing
        [bool]$includeDateInFileName, # Flag to include date in CSV file name
        [bool]$includeSummary, # Flag to display summary information
        [bool]$silentMode = $False,  # Flag to suppress prompts
        [bool]$sendMail  # Flag to send email                          
    )

    if ([string]::IsNullOrEmpty($folderPath)) {
        $folderPath = "$PSScriptRoot/Page Counter"
    }

    $dataArray = @()  # Array to store document data
    $invalidFiles = @()  # Array to store files with issues

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

    # Process Word files
    foreach ($wordFile in $wordFiles) {
        $wordPageCount = Get-WordDocPageCount -filePath $wordFile.FullName -wordApp ([ref]$word)
        if ($wordPageCount -eq 0) {
            $invalidFiles += $wordFile.Name
        } else {
            $dataArray += [pscustomobject]@{
                FileName = $wordFile.Name
                Pages    = $wordPageCount
                Type     = "Word Document"
            }
        }
    }

    $word.Quit()

    # Process PDF files
    foreach ($pdfFile in $pdfFiles) {
        $pdfPageCount = Get-PDFPageCount -filePath $pdfFile.FullName
        if ($pdfPageCount -eq 0) {
            $invalidFiles += $pdfFile.Name
        } else {
            $dataArray += [pscustomobject]@{
                FileName = $pdfFile.Name
                Pages    = $pdfPageCount
                Type     = "PDF Document"
            }
        }
    }

    # Display invalid files if any
    if ($invalidFiles.Count -gt 0 -and -not $silentMode) {
        Write-Host "`nFiles with issues:"
        $invalidFiles | ForEach-Object { Write-Host " - $_" }
    }

    $totalWordPages = ($dataArray | Where-Object { $_.Type -eq 'Word Document' } | Measure-Object -Property Pages -Sum).Sum
    $totalPdfPages = ($dataArray | Where-Object { $_.Type -eq 'PDF Document' } | Measure-Object -Property Pages -Sum).Sum
    
    if ($includeSummary -and -not $silentMode) {
        # Display summary information if flag is set    
        Write-Host "`n-------------------------------------------------Summary------------------------------------------------"
        Write-Host "Total Word Pages: $totalWordPages"
        Write-Host "Total PDF Pages: $totalPdfPages"
        Write-Host "Total Word files processed: $($wordFiles.Count)"
        Write-Host "Total PDF files processed: $($pdfFiles.Count)"
        Write-Host "Total files processed: $($wordFiles.Count + $pdfFiles.Count)"
        Write-Host "Total Excluded files: $($invalidFiles.Count)"   
        Write-Host "--------------------------------------------------------------------------------------------------------"
    }

    # Handle CSV file writing and email functionality
    if ($writeCSV) {
        csvMenu -dataArray $dataArray -csvPathFull $csvPathFull -silentMode $silentMode
    }
    if ($sendMail -and $toEmailAddress) {
        Send-Email -toEmailAddress $toEmailAddress -dataArray $dataArray -invalidFiles $invalidFiles
    }

    Write-Host "Operation Completed Successfully"
}

function csvMenu {
    param (
        [array]$dataArray,
        [string]$csvPathFull,
        [bool]$silentMode
    )
    
    $csvPath = "$PSScriptRoot/"

    # Generate CSV file path with current date if needed
    $csvPathFull = if ($includeDateInFileName) {
        $csvPath + "document_page_counts_" + (Get-Date -Format "yyyyMMdd") + ".csv"
    }
    else {
        $csvPath + "document_page_counts.csv"
    }

    # Recipient email address
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

    foreach ($data in $dataArray) {
        # Format data as CSV line and append to file
        "$($data.FileName),$($data.Pages),$($data.Type)" | Out-File -Append $csvPathFull
    }
}

# Function to send email with attachment
function Send-Email {
    param (
        [string]$toEmailAddress,
        [array]$dataArray,  # Array of objects containing file data
        [array]$invalidFiles
    )

    $totalPages = $dataArray.Pages.Count

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

    $emailBody += "The following files had issues and were excluded:`n"
    foreach ($file in $invalidFiles) {
        $emailBody += " - $file`n"
    }
    
    # Set the email properties
    $mail.Subject = "Page Counter Results - $currentDateTime"
    $mail.Body = $emailBody
    $mail.To = $toEmailAddress

    # Send the email
    $mail.Send()
    Write-Host "Email sent to: $toEmailAddress"
}

# Function to save the email address in JSON format
function Save-EmailConfig {
    param (
        [string]$emailAddress
    )

    # Read existing configuration
    $configData = @{}
    if (Test-Path $configFilePath) {
        $configData = Get-Content -Path $configFilePath | ConvertFrom-Json
    }

    # Update only the EmailAddress field
    $configData.EmailAddress = $emailAddress

    # Convert to JSON and save back
    $configData | ConvertTo-Json -Depth 3 | Set-Content -Path $configFilePath -Force
}

# Function to get the saved email address from the JSON file
function Get-SavedEmailAddress {
    if (Test-Path $configFilePath) {
        $configData = Get-Content -Path $configFilePath | ConvertFrom-Json
        return $configData.EmailAddress
    }
    return $null
}

# Function to create the GUI for entering the email address
function Show-EmailConfigForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Configure Email Address"
    $form.Size = New-Object System.Drawing.Size(400, 200)
    $form.StartPosition = "CenterScreen"

    # Label
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Enter the email address to send reports to:"
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $form.Controls.Add($label)

    # Textbox for email address
    $emailTextbox = New-Object System.Windows.Forms.TextBox
    $emailTextbox.Size = New-Object System.Drawing.Size(300, 20)
    $emailTextbox.Location = New-Object System.Drawing.Point(20, 60)
    $form.Controls.Add($emailTextbox)

    # Save button
    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Save"
    $saveButton.Location = New-Object System.Drawing.Point(150, 100)
    $saveButton.Add_Click({
        Save-EmailConfig -emailAddress $emailTextbox.Text
        [System.Windows.Forms.MessageBox]::Show("Email address saved successfully.", "Saved", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $form.Close()
    })
    $form.Controls.Add($saveButton)

    # Show the form
    $form.ShowDialog()
}

# Main logic to check and retrieve or prompt for the email address
$emailAddress = Get-SavedEmailAddress
if (-not $emailAddress) {
    Show-EmailConfigForm 
    $emailAddress = Get-SavedEmailAddress  
}

$configData = Get-Content -Path $configFilePath | ConvertFrom-Json

Get-FolderPageCounts -folderPath $configData.FolderPath -writeCSV $configData.WriteCSV `
    -includeDateInFileName $configData.IncludeDateInFileName -includeSummary $configData.IncludeSummary `
    -silentMode $configData.SilentMode -toEmailAddress $configData.EmailAddress -sendMail $configData.SendMail