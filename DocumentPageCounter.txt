function Get-WordDocPageCount {
    param (
        [string]$filePath
    )

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Open($filePath)
    $pageCount = $doc.ComputeStatistics([Microsoft.Office.Interop.Word.WdStatistic]::wdStatisticPages)
    $doc.Close()
    $word.Quit()
    return $pageCount
}

function Get-PDFPageCount {
    param (
        [string]$filePath
    )

    $content = Get-Content -Path $filePath -Raw
    $matches = [regex]::Matches($content, "/Type\s*/Page[^s]")
    return $matches.Count
}

function Get-FolderPageCounts {
    param (
        [string]$folderPath
    )

    $wordFiles = Get-ChildItem -Path $folderPath -Filter "*.docx" -Recurse
    $pdfFiles = Get-ChildItem -Path $folderPath -Filter "*.pdf" -Recurse

    foreach ($wordFile in $wordFiles) {
        $wordPageCount = Get-WordDocPageCount $wordFile.FullName
        Write-Host "File: $($wordFile.Name) - Pages: $wordPageCount - Type: Word Document"
    }

    foreach ($pdfFile in $pdfFiles) {
        $pdfPageCount = Get-PDFPageCount $pdfFile.FullName
        Write-Host "File: $($pdfFile.Name) - Pages: $pdfPageCount - Type: PDF Document"
    }
}

$folderPath = "./Documents"
Get-FolderPageCounts $folderPath