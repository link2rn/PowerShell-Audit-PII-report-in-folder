# Define the directory and file types to search
$SearchPath = "D:\Working\CodeLab\Powershell\CodeLab\PowerShell\Report PII\SampleFiles"
$FileTypes = "*.txt", "*.csv", "*.docx", "*.doc", "*.xlsx", "*.xls", "*.pptx", "*.ppt", "*.rtf"

# Define patterns for PII (e.g., SSNs, Credit Card Numbers, etc.)
$Patterns = @(
    @{ Name = "CreditCard_Visa"; Regex = '\b4[0-9]{3}[-\s]?[0-9]{4}[-\s]?[0-9]{4}[-\s]?[0-9]{4}\b' },
    @{ Name = "CreditCard_MasterCard"; Regex = '\b5[1-5][0-9]{2}[-\s]?[0-9]{4}[-\s]?[0-9]{4}[-\s]?[0-9]{4}\b' },
    @{ Name = "CreditCard_Amex"; Regex = '\b3[47][0-9]{2}[-\s]?[0-9]{6}[-\s]?[0-9]{5}\b' },
    @{ Name = "CreditCard_DinersClub"; Regex = '\b(3(0[0-5]|[68][0-9]))[-\s]?[0-9]{4}[-\s]?[0-9]{4}[-\s]?[0-9]{4}\b' },
    @{ Name = "CreditCard_Discover"; Regex = '\b6(?:011|5[0-9]{2})[-\s]?[0-9]{4}[-\s]?[0-9]{4}[-\s]?[0-9]{4}\b' },
    @{ Name = "CreditCard_JCB"; Regex = '\b(?:2131|1800|35\d{3})[-\s]?[0-9]{4}[-\s]?[0-9]{4}[-\s]?[0-9]{4}\b' },
    @{ Name = "SSN"; Regex = '\b\d{3}[-\s]?\d{2}[-\s]?\d{4}\b' },
    @{ Name = "Email"; Regex = '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b' },
    @{ Name = "PhoneNumber"; Regex = '\b\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}\b' }
)

# Collect results
$Results = @()

# Helper function to read document content
function Get-DocumentContent {
    param (
        [string]$FilePath,
        [string]$Extension
    )
    switch ($Extension) {
        {($_ -eq ".docx") -or ($_ -eq ".doc") -or ($_ -eq ".rtf")} {
            $Word = New-Object -ComObject Word.Application
            $Word.Visible = $false
            $Document = $Word.Documents.Open($FilePath, [ref]$false, [ref]$true, [ref]$false)
            $Content = $Document.Content.Text -split "`r`n"
            $Document.Close()
            $Word.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Document) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            return $Content
        }
        {($_ -eq ".xlsx") -or ($_ -eq ".xls")} {
            $Excel = New-Object -ComObject Excel.Application
            $Excel.Visible = $false
            $Workbook = $Excel.Workbooks.Open($FilePath, 0, $true, 1, "", "", $true, [System.Type]::Missing, [System.Type]::Missing, $true, [System.Type]::Missing, [System.Type]::Missing, [System.Type]::Missing, [System.Type]::Missing, [System.Type]::Missing)
            $Content = @()
            foreach ($Worksheet in $Workbook.Worksheets) {
                $Range = $Worksheet.UsedRange
                foreach ($Row in $Range.Rows) {
                    $Content += ($Row.Text -split "`r`n")
                }
            }
            $Workbook.Close($false)
            $Excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
            return $Content
        }
        {($_ -eq ".pptx") -or ($_ -eq ".ppt")} {
            $PowerPoint = New-Object -ComObject PowerPoint.Application
            $PowerPoint.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
            $Presentation = $PowerPoint.Presentations.Open($FilePath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
            $Content = @()
            foreach ($Slide in $Presentation.Slides) {
                if ($Slide.NotesPage.Shapes.Count -ge 2) {
                    $Content += ($Slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text -split "`r`n")
                }
            }
            $Presentation.Close()
            $PowerPoint.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($PowerPoint) | Out-Null
            return $Content
        }
        default {
            return Get-Content -Path $FilePath -ErrorAction SilentlyContinue
        }
    }
}

# Search for PII patterns in files
$TotalFiles = 0
$ProcessedFiles = 0
foreach ($FileType in $FileTypes) {
    $Files = Get-ChildItem -Path $SearchPath -Filter $FileType -Recurse -ErrorAction SilentlyContinue
    $TotalFiles += $Files.Count
}

Write-Progress -Activity "Scanning Files for PII" -Status "Initializing..." -PercentComplete 0

foreach ($FileType in $FileTypes) {
    $Files = Get-ChildItem -Path $SearchPath -Filter $FileType -Recurse -ErrorAction SilentlyContinue
    foreach ($File in $Files) {
        $Content = Get-DocumentContent -FilePath $File.FullName -Extension $File.Extension
        $TotalLines = $Content.Length
        $ProcessedLines = 0

        for ($LineNumber = 0; $LineNumber -lt $TotalLines; $LineNumber++) {
            $Line = $Content[$LineNumber]
            foreach ($Pattern in $Patterns) {
                if ($Line -match $Pattern.Regex) {
                    $Results += [PSCustomObject]@{
                        FileName = $File.Name
                        FilePath = $File.FullName
                        DirectoryName = $File.DirectoryName
                        PatternName = $Pattern.Name
                        MatchCount = 1
                        LineNumber = $LineNumber + 1
                    }
                }
            }
            $ProcessedLines++
            $PercentComplete = [math]::Round(($ProcessedLines / $TotalLines) * 100)
            Write-Progress -Activity "Scanning File: $($File.Name)" -Status "Processing lines..." -PercentComplete $PercentComplete
        }
        $ProcessedFiles++
        $PercentComplete = [math]::Round(($ProcessedFiles / $TotalFiles) * 100)
        Write-Progress -Activity "Scanning Files for PII" -Status "Processing file $($File.Name)" -PercentComplete $PercentComplete
    }
}

# Generate HTML report
$Html = @"
<html>
<head>
    <title>Report of Files Containing PII</title>
    <style>
        table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
        table, th, td { border: 1px solid black; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        h3 { margin-bottom: 5px; }
    </style>
</head>
<body>
    <h2>Report of Files Containing PII</h2>
"@

$GroupedResults = $Results | Group-Object FilePath

foreach ($Group in $GroupedResults) {
    $FilePath = $Group.Name
    $Html += "<h3>File: $($Group.Group[0].FileName)</h3>"
    $Html += "<p><strong>Full Path:</strong> $FilePath</p>"
    $Html += "<table>
                <tr>
                    <th>Pattern</th>
                    <th>Match Count</th>
                    <th>Line Number</th>
                </tr>"
    foreach ($Result in $Group.Group) {
        $Html += "<tr><td>$($Result.PatternName)</td><td>$($Result.MatchCount)</td><td>$($Result.LineNumber)</td></tr>"
    }
    $Html += "</table>"
}

$Html += @"
</body>
</html>
"@

# Output HTML report
$ReportPath = "$SearchPath\PII_Report.html"
$Html | Out-File -FilePath $ReportPath -Encoding UTF8

Write-Output "Report generated at: $ReportPath"
