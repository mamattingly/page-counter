# function Get-WordDocPageCount {
#     param (
#         [string]$filePath
#     )

#     $word = New-Object -ComObject Word.Application
#     $word.Visible = $false
#     $doc = $word.Documents.Open($filePath)
#     $pageCount = $doc.ComputeStatistics([Microsoft.Office.Interop.Word.WdStatistic]::wdStatisticPages)
#     $doc.Close()
#     $word.Quit()
#     return $pageCount
# }

# function Get-PDFPageCount {
#     param (
#         [string]$filePath
#     )

#     $content = Get-Content -Path $filePath -Raw
#     $matches = [regex]::Matches($content, "/Type\s*/Page[^s]")
#     return $matches.Count
# }

# function Get-FolderPageCounts {
#     param (
#         [string]$folderPath
#     )

#     if ($folderPath -eq "./Documents") {
#         $folderPath = Read-Host "Drag a folder or leave blank for the default folder"
#     }

#     $wordFiles = Get-ChildItem -Path $folderPath -Filter "*.docx" -Recurse
#     $pdfFiles = Get-ChildItem -Path $folderPath -Filter "*.pdf" -Recurse

#     foreach ($wordFile in $wordFiles) {
#         $wordPageCount = Get-WordDocPageCount $wordFile.FullName
#         Write-Host "File: $($wordFile.Name) - Pages: $wordPageCount - Type: Word Document"
#     }

#     foreach ($pdfFile in $pdfFiles) {
#         $pdfPageCount = Get-PDFPageCount $pdfFile.FullName
#         Write-Host "File: $($pdfFile.Name) - Pages: $pdfPageCount - Type: PDF Document"
#     }
# }

# Get-FolderPageCounts -folderPath "./Documents"


function Get-WordDocPageCount {
    param (
        [string]$filePath,
        [ref]$wordApp
    )

    try {
        $doc = $wordApp.Value.Documents.Open($filePath)
        if ($null -eq $doc) {
            Write-Host "Error: Unable to open Word document at path: $filePath"
            return 0
        }

        $pageCount = $doc.ComputeStatistics([Microsoft.Office.Interop.Word.WdStatistic]::wdStatisticPages)
        $doc.Close()
        return $pageCount
    } catch {
        Write-Host "Error processing Word document: $_"
        return 0
    }
}

function Get-PDFPageCount {
    param (
        [string]$filePath
    )

    if ([string]::IsNullOrEmpty($filePath)) {
        Write-Host "Error: File path is empty or null."
        return 0
    }

    try {
        $content = Get-Content -Path $filePath -Raw -ReadCount 0
        $matches = [regex]::Matches($content, "/Type\s*/Page[^s]")
        return $matches.Count
    } catch {
        Write-Host "Error processing PDF document: $_"
        return 0
    }
}

# function Get-FolderPageCounts {
#     param (
#         [string]$folderPath
#     )

#     if ($folderPath -eq "./Documents") {
#         $folderPath = Read-Host "Drag a folder or leave blank for the default folder"
#     }

#     $wordFiles = Get-ChildItem -Path $folderPath -Filter "*.docx" -Recurse -ErrorAction SilentlyContinue
#     $pdfFiles = Get-ChildItem -Path $folderPath -Filter "*.pdf" -Recurse -ErrorAction SilentlyContinue

#     $word = New-Object -ComObject Word.Application
#     $word.Visible = $false

#     foreach ($wordFile in $wordFiles) {
#         $wordPageCount = Get-WordDocPageCount -filePath $wordFile.FullName -wordApp ([ref]$word)
#         Write-Host "File: $($wordFile.Name) - Pages: $wordPageCount - Type: Word Document"
#     }

#     $word.Quit()

#     foreach ($pdfFile in $pdfFiles) {
#         $pdfPageCount = Get-PDFPageCount -filePath $pdfFile.FullName
#         Write-Host "File: $($pdfFile.Name) - Pages: $pdfPageCount - Type: PDF Document"
#     }
# }
function Add-DataToCSV {
    param (
        [string]$fileName,
        [string]$pages,
        [string]$type,
        [string]$csvPath
    )

    $data = "$fileName,$pages,$type"
    $data | Out-File -Append $csvPath
}

function Get-FolderPageCounts {
    param (
        [string]$folderPath
    )

    if (-not $folderPath) {
        $folderPath = Read-Host "Drag a folder or leave blank for the default folder"
        if (-not $folderPath) {
            $folderPath = "./Documents"
        }
    }

    $wordFiles = Get-ChildItem -Path $folderPath -Filter "*.docx" -Recurse -ErrorAction SilentlyContinue
    $pdfFiles = Get-ChildItem -Path $folderPath -Filter "*.pdf" -Recurse -ErrorAction SilentlyContinue

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false

    foreach ($wordFile in $wordFiles) {
        $wordPageCount = Get-WordDocPageCount -filePath $wordFile.FullName -wordApp ([ref]$word)
        Write-Host "File: $($wordFile.Name) - Pages: $wordPageCount - Type: Word Document"
    }
    $word.Quit()

    foreach ($pdfFile in $pdfFiles) {
        $pdfPageCount = Get-PDFPageCount -filePath $pdfFile.FullName
        Write-Host "File: $($pdfFile.Name) - Pages: $pdfPageCount - Type: PDF Document"
    }

    # Prompt for CSV handling after processing files
    $csvPath = "$env:USERPROFILE\Downloads\document_page_counts.csv"

    if (Test-Path $csvPath) {
        $choice = Read-Host "The CSV file already exists. Choose an option:`n1. Append`n2. Overwrite`n3. Cancel"

        switch ($choice) {
            "1" {
                # Append mode, no need to add header
            }
            "2" {
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
    } else {
        $csvData = "File,Pages,Type"
        $csvData | Out-File $csvPath
    }

    # Adding data to the CSV after determining the mode
    foreach ($wordFile in $wordFiles) {
        $wordPageCount = Get-WordDocPageCount -filePath $wordFile.FullName -wordApp ([ref]$word)
        Add-DataToCSV -fileName $wordFile.Name -pages $wordPageCount -type "Word Document" -csvPath $csvPath
    }

    foreach ($pdfFile in $pdfFiles) {
        $pdfPageCount = Get-PDFPageCount -filePath $pdfFile.FullName
        Add-DataToCSV -fileName $pdfFile.Name -pages $pdfPageCount -type "PDF Document" -csvPath $csvPath
    }

    Write-Host "CSV file updated at: $csvPath"
}

Get-FolderPageCounts
