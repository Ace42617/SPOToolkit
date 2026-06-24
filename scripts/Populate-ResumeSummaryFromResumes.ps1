<#
.SYNOPSIS
    Scans a folder of Word resumes, extracts key fields, and fills a Word template for each.

.DESCRIPTION
    Uses Microsoft Word COM automation to read .doc/.docx resume files, extract common
    contact and section data (email, phone, skills, experience, etc.), and produce one
    populated document per resume from a template.

    The template should contain placeholders in double curly braces, for example:
      {{FullName}}, {{Email}}, {{Phone}}, {{LinkedIn}}, {{Summary}},
      {{Skills}}, {{Experience}}, {{Education}}, {{SourceFileName}}

    Extraction uses regex and section-header heuristics. Resume layouts vary widely, so
    review output for accuracy and adjust patterns in the script if your resumes follow
    a consistent format.

.PARAMETER ResumeFolder
    Folder containing source resume Word files.

.PARAMETER TemplatePath
    Path to the Word template (.docx or .dotx) with {{Placeholder}} tokens.

.PARAMETER OutputFolder
    Folder where populated documents are written. Created if missing.

.PARAMETER FilePattern
    File filter for resumes. Default: *.docx

.PARAMETER Overwrite
    Replace existing output files. Default: false.

.PARAMETER KeepWordOpen
    Leave the Word application running after the script finishes (useful for debugging).

.EXAMPLE
    .\Populate-ResumeSummaryFromResumes.ps1 `
        -ResumeFolder 'C:\Resumes\incoming' `
        -TemplatePath 'C:\Templates\CandidateSummary.docx' `
        -OutputFolder 'C:\Resumes\summaries'

.NOTES
    Requires: Microsoft Word installed on Windows.
    Word may prompt about macros or protected view depending on your trust settings.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string] $ResumeFolder,

    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string] $TemplatePath,

    [Parameter(Mandatory = $true)]
    [string] $OutputFolder,

    [string] $FilePattern = '*.docx',

    [switch] $Overwrite,

    [switch] $KeepWordOpen
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function New-WordApplication {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0  # wdAlertsNone
    return $word
}

function Close-WordApplication {
    param(
        [Parameter(Mandatory = $true)]
        $WordApp,

        [switch] $ForceQuit
    )

    if (-not $WordApp) { return }

    if ($ForceQuit -or -not $KeepWordOpen) {
        [void] [System.Runtime.InteropServices.Marshal]::ReleaseComObject($WordApp)
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        $WordApp.Quit()
    }
}

function Get-WordDocumentPlainText {
    param(
        [Parameter(Mandatory = $true)]
        $WordApp,

        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    $document = $null
    try {
        $document = $WordApp.Documents.Open(
            $Path,
            $false,  # ConfirmConversions
            $true,   # ReadOnly
            $false,  # AddToRecentFiles
            '',      # PasswordDocument
            '',      # PasswordTemplate
            $false,  # Revert
            '',      # WritePasswordDocument
            '',      # WritePasswordTemplate
            0,       # Format (wdOpenFormatAuto)
            $null,   # Encoding
            $false,  # Visible
            $false,  # OpenAndRepair
            0,       # DocumentDirection
            $false   # NoEncodingDialog
        )

        return [string] $document.Content.Text
    }
    finally {
        if ($document) {
            $document.Close($false)
            [void] [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document)
        }
    }
}

function Get-FirstMatch {
    param(
        [string] $Text,
        [string] $Pattern
    )

    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $match = [regex]::Match($Text, $Pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($match.Success) {
        return $match.Value.Trim()
    }
    return $null
}

function Get-SectionText {
    param(
        [string[]] $Lines,
        [string[]] $HeaderPatterns,
        [string[]] $StopHeaderPatterns
    )

    $startIndex = -1
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $line = $Lines[$i].Trim()
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        foreach ($pattern in $HeaderPatterns) {
            if ($line -match $pattern) {
                $startIndex = $i + 1
                break
            }
        }
        if ($startIndex -ge 0) { break }
    }

    if ($startIndex -lt 0) { return $null }

    $sectionLines = New-Object System.Collections.Generic.List[string]
    for ($j = $startIndex; $j -lt $Lines.Count; $j++) {
        $line = $Lines[$j].Trim()
        if ([string]::IsNullOrWhiteSpace($line)) {
            if ($sectionLines.Count -gt 0) { break }
            continue
        }

        $isStop = $false
        foreach ($stopPattern in $StopHeaderPatterns) {
            if ($line -match $stopPattern) {
                $isStop = $true
                break
            }
        }
        if ($isStop) { break }

        $sectionLines.Add($line) | Out-Null
    }

    if ($sectionLines.Count -eq 0) { return $null }
    return ($sectionLines -join [Environment]::NewLine).Trim()
}

function Get-ProbableFullName {
    param(
        [string[]] $Lines,
        [string] $Email
    )

    foreach ($line in $Lines) {
        $candidate = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        if ($candidate.Length -gt 80) { continue }
        if ($candidate -match '@|https?://|linkedin\.com|phone|email|resume|curriculum vitae|\d{3}[-.\s]?\d{3}') {
            continue
        }
        if ($candidate -match '^[A-Za-z][A-Za-z''.\- ]{1,60}$') {
            return $candidate
        }
    }

    if ($Email) {
        $localPart = ($Email -split '@', 2)[0]
        $nameGuess = ($localPart -replace '[._\-+]', ' ').Trim()
        if ($nameGuess -match '^[A-Za-z''.\- ]+$') {
            return (Get-Culture).TextInfo.ToTitleCase($nameGuess.ToLower())
        }
    }

    return $null
}

function ConvertFrom-ResumeText {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Text,

        [Parameter(Mandatory = $true)]
        [string] $SourceFileName
    )

    $normalized = $Text -replace "`r`n|`r|`n", [Environment]::NewLine
    $normalized = $normalized -replace "`v|`f", [Environment]::NewLine
    $lines = $normalized -split [Environment]::NewLine

    $email = Get-FirstMatch $normalized '(?<!\w)[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}(?!\w)'
    $phone = Get-FirstMatch $normalized '(?<!\w)(?:\+?\d{1,3}[\s\-.]?)?(?:\(?\d{2,4}\)?[\s\-.]?)?\d{3}[\s\-.]?\d{3,4}(?:[\s\-.]?\d{1,5})?(?!\w)'
    $linkedIn = Get-FirstMatch $normalized '(?:https?://)?(?:www\.)?linkedin\.com/in/[A-Za-z0-9\-_%]+/?'

    $sectionStopPatterns = @(
        '^(?:summary|profile|objective|professional summary)\b'
        '^(?:skills|technical skills|core competencies|competencies)\b'
        '^(?:experience|work experience|employment|professional experience|work history|career history)\b'
        '^(?:education|academic background|qualifications)\b'
        '^(?:certifications?|licenses?|projects?|awards?|references?)\b'
    )

    $summary = Get-SectionText -Lines $lines -HeaderPatterns @(
        '^(?:summary|profile|objective|professional summary)\s*:?\s*$'
        '^(?:summary|profile|objective|professional summary)\b'
    ) -StopHeaderPatterns ($sectionStopPatterns | Where-Object { $_ -notmatch 'summary|profile|objective' })

    $skills = Get-SectionText -Lines $lines -HeaderPatterns @(
        '^(?:skills|technical skills|core competencies|competencies)\s*:?\s*$'
        '^(?:skills|technical skills|core competencies|competencies)\b'
    ) -StopHeaderPatterns ($sectionStopPatterns | Where-Object { $_ -notmatch 'skills|competencies' })

    $experience = Get-SectionText -Lines $lines -HeaderPatterns @(
        '^(?:experience|work experience|employment|professional experience|work history|career history)\s*:?\s*$'
        '^(?:experience|work experience|employment|professional experience|work history|career history)\b'
    ) -StopHeaderPatterns ($sectionStopPatterns | Where-Object { $_ -notmatch 'experience|employment|work history|career history' })

    $education = Get-SectionText -Lines $lines -HeaderPatterns @(
        '^(?:education|academic background|qualifications)\s*:?\s*$'
        '^(?:education|academic background|qualifications)\b'
    ) -StopHeaderPatterns ($sectionStopPatterns | Where-Object { $_ -notmatch 'education|qualifications|academic' })

    $fullName = Get-ProbableFullName -Lines $lines -Email $email

    return [ordered]@{
        FullName       = if ($fullName) { $fullName } else { '' }
        Email          = if ($email) { $email } else { '' }
        Phone          = if ($phone) { $phone } else { '' }
        LinkedIn       = if ($linkedIn) { $linkedIn } else { '' }
        Summary        = if ($summary) { $summary } else { '' }
        Skills         = if ($skills) { $skills } else { '' }
        Experience     = if ($experience) { $experience } else { '' }
        Education      = if ($education) { $education } else { '' }
        SourceFileName = $SourceFileName
    }
}

function Set-WordTemplatePlaceholders {
    param(
        [Parameter(Mandatory = $true)]
        $WordApp,

        [Parameter(Mandatory = $true)]
        [string] $TemplatePath,

        [Parameter(Mandatory = $true)]
        [string] $OutputPath,

        [Parameter(Mandatory = $true)]
        [hashtable] $Data
    )

    $document = $null
    try {
        $document = $WordApp.Documents.Open(
            $TemplatePath,
            $false,
            $false,
            $false,
            '',
            '',
            $false,
            '',
            '',
            0,
            $null,
            $false,
            $false,
            0,
            $false
        )

        foreach ($key in $Data.Keys) {
            $placeholder = "{{$key}}"
            $value = [string] $Data[$key]
            if ($null -eq $value) { $value = '' }

            $find = $document.Content.Find
            $find.ClearFormatting()
            $find.Replacement.ClearFormatting()
            $find.Text = $placeholder
            $find.Replacement.Text = $value
            $find.Forward = $true
            $find.Wrap = 1  # wdFindContinue
            $find.Format = $false
            $find.MatchCase = $true
            $find.MatchWholeWord = $false
            $find.MatchWildcards = $false

            [void] $find.Execute(
                $find.Text,
                $false,
                $false,
                $false,
                $false,
                $false,
                $true,
                1,
                $false,
                $find.Replacement.Text,
                2  # wdReplaceAll
            )
        }

        $outputDirectory = Split-Path -LiteralPath $OutputPath -Parent
        if (-not (Test-Path -LiteralPath $outputDirectory)) {
            New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
        }

        $document.SaveAs2([ref] $OutputPath)
    }
    finally {
        if ($document) {
            $document.Close($false)
            [void] [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document)
        }
    }
}

function Get-SafeOutputPath {
    param(
        [string] $OutputFolder,
        [string] $BaseName
    )

    $safeName = ($BaseName -replace '[\\/:*?"<>|]', '_').Trim()
    if ([string]::IsNullOrWhiteSpace($safeName)) {
        $safeName = 'ResumeSummary'
    }

    return Join-Path $OutputFolder ("{0}_Summary.docx" -f $safeName)
}

if (-not (Test-Path -LiteralPath $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

$resumeFiles = @(Get-ChildItem -LiteralPath $ResumeFolder -Filter $FilePattern -File | Sort-Object Name)
if ($resumeFiles.Count -eq 0) {
    Write-Warning "No files matching '$FilePattern' were found in '$ResumeFolder'."
    return
}

Write-Host "Found $($resumeFiles.Count) resume file(s) in '$ResumeFolder'."

$wordApp = $null
$results = New-Object System.Collections.Generic.List[object]

try {
    $wordApp = New-WordApplication

    foreach ($resumeFile in $resumeFiles) {
        $outputPath = Get-SafeOutputPath -OutputFolder $OutputFolder -BaseName $resumeFile.BaseName
        if ((Test-Path -LiteralPath $outputPath) -and -not $Overwrite) {
            Write-Warning "Skipping '$($resumeFile.Name)' because output already exists: $outputPath"
            continue
        }

        if ($PSCmdlet.ShouldProcess($resumeFile.FullName, "Extract and populate template")) {
            Write-Host "Processing: $($resumeFile.Name)"

            $plainText = Get-WordDocumentPlainText -WordApp $wordApp -Path $resumeFile.FullName
            $data = ConvertFrom-ResumeText -Text $plainText -SourceFileName $resumeFile.Name

            Set-WordTemplatePlaceholders `
                -WordApp $wordApp `
                -TemplatePath $TemplatePath `
                -OutputPath $outputPath `
                -Data $data

            $results.Add([pscustomobject]@{
                SourceFile = $resumeFile.Name
                OutputFile = Split-Path -Leaf $outputPath
                FullName   = $data.FullName
                Email      = $data.Email
                Phone      = $data.Phone
            }) | Out-Null
        }
    }
}
finally {
    Close-WordApplication -WordApp $wordApp
}

if ($results.Count -gt 0) {
    Write-Host ""
    Write-Host "Completed $($results.Count) document(s):"
    $results | Format-Table -AutoSize
}
else {
    Write-Host "No documents were generated."
}
