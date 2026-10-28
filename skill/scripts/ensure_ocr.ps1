[CmdletBinding()]
param()

# Install and configure Tesseract OCR with the Traditional Chinese model.
# Run this once before OCR-ing scanned PDFs.
# Requires PowerShell 7 (pwsh) on both Windows and macOS.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

if ($null -eq $IsWindows) {
    throw 'PowerShell 7 (pwsh) is required; 5.1 has no $IsWindows and would silently take the wrong branch.'
}

function Find-Tesseract {
    $command = Get-Command 'tesseract' -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }

    if (-not $IsWindows) { return $null }   # macOS: PATH is the only place Homebrew puts it

    $candidates = @(
        (Join-Path ([Environment]::GetFolderPath('ProgramFiles')) 'Tesseract-OCR' 'tesseract.exe'),
        (Join-Path ([Environment]::GetFolderPath('ProgramFilesX86')) 'Tesseract-OCR' 'tesseract.exe')
    )
    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path -LiteralPath $candidate)) { return $candidate }
    }
    return $null
}

if (-not $IsWindows) {
    # ---- macOS ----------------------------------------------------------------
    # Much shorter than the Windows path: Homebrew's tessdata directory is writable
    # by the user, so none of the staging below is needed, and tesseract-lang ships
    # chi_tra directly. ocrmypdf also needs ghostscript, which Windows handles elsewhere.
    Write-Host '[1/2] Checking Tesseract OCR...'
    $TesseractPath = Find-Tesseract
    if (-not $TesseractPath) {
        $Brew = Get-Command 'brew' -ErrorAction SilentlyContinue
        if (-not $Brew) {
            throw 'Tesseract OCR and Homebrew were not found. Install Homebrew from https://brew.sh and run this script again.'
        }
        Write-Host 'Installing tesseract, tesseract-lang and ghostscript with Homebrew...'
        & $Brew.Source install tesseract tesseract-lang ghostscript
        $TesseractPath = Find-Tesseract
        if (-not $TesseractPath) {
            throw 'Tesseract OCR was installed, but its executable was not found. Restart the Agent and run this script again.'
        }
    }
    Write-Host "Tesseract OCR: $TesseractPath"

    Write-Host '[2/2] Verifying the Traditional Chinese model...'
    $Languages = & $TesseractPath --list-langs 2>&1
    if ($Languages -notcontains 'chi_tra') {
        $Brew = Get-Command 'brew' -ErrorAction SilentlyContinue
        if ($Brew) {
            Write-Host 'chi_tra missing; installing tesseract-lang...'
            & $Brew.Source install tesseract-lang
            $Languages = & $TesseractPath --list-langs 2>&1
        }
    }
    if ($Languages -notcontains 'chi_tra') {
        throw 'chi_tra is still missing. Install it with: brew install tesseract-lang'
    }

    if (-not (Get-Command 'gs' -ErrorAction SilentlyContinue)) {
        Write-Host 'Warning: ghostscript not found. ocrmypdf needs it: brew install ghostscript'
    }

    Write-Host 'OCR is ready. Use lang chi_tra+eng for Traditional Chinese handouts.'
    exit 0
}

# ---- Windows ------------------------------------------------------------------
Write-Host '[1/3] Checking Tesseract OCR...'
$TesseractPath = Find-Tesseract
if (-not $TesseractPath) {
    $WinGet = Get-Command 'winget' -ErrorAction SilentlyContinue
    if (-not $WinGet) {
        throw 'Tesseract OCR and WinGet were not found. Install UB-Mannheim.TesseractOCR and run this script again.'
    }
    Write-Host 'Installing Tesseract OCR with WinGet...'
    & $WinGet.Source install --exact --id UB-Mannheim.TesseractOCR --silent --accept-package-agreements --accept-source-agreements
    if ($LASTEXITCODE -ne 0) {
        throw "Tesseract OCR installation failed. WinGet exit code: $LASTEXITCODE"
    }
    $TesseractPath = Find-Tesseract
    if (-not $TesseractPath) {
        throw 'Tesseract OCR was installed, but its executable was not found. Restart the Agent and run this script again.'
    }
}
Write-Host "Tesseract OCR: $TesseractPath"

# Models cannot be written into Program Files, so keep them under the user profile, next to the
# shared environment. Not %LOCALAPPDATA%: run from the Claude desktop app (an MSIX package) that
# write lands in its private LocalCache, while TESSDATA_PREFIX is seen by every program.
Write-Host '[2/3] Configuring English, orientation, and Traditional Chinese OCR data...'
$SourceTessData = Join-Path (Split-Path -Parent $TesseractPath) 'tessdata'
$TessData = Join-Path $HOME '.local' 'share' 'file-toolkit' 'tessdata'
New-Item -ItemType Directory -Path $TessData -Force | Out-Null

foreach ($fileName in @('eng.traineddata', 'osd.traineddata', 'pdf.ttf')) {
    $sourceFile = Join-Path $SourceTessData $fileName
    if (-not (Test-Path -LiteralPath $sourceFile)) {
        throw "Missing Tesseract data file: $sourceFile"
    }
    Copy-Item -LiteralPath $sourceFile -Destination $TessData -Force
}

foreach ($folderName in @('configs', 'tessconfigs')) {
    $sourceFolder = Join-Path $SourceTessData $folderName
    $targetFolder = Join-Path $TessData $folderName
    if (-not (Test-Path -LiteralPath $sourceFolder)) {
        throw "Missing Tesseract configuration folder: $sourceFolder"
    }
    New-Item -ItemType Directory -Path $targetFolder -Force | Out-Null
    Get-ChildItem -LiteralPath $sourceFolder -File | Copy-Item -Destination $targetFolder -Force
}

$ChiTraPath = Join-Path $TessData 'chi_tra.traineddata'
$ChiTraHash = '529C5B5797D64B126065CD55F2BB4C7FD7B15790798091B1FF259941A829330B'
$ChiTraUrl = 'https://raw.githubusercontent.com/tesseract-ocr/tessdata_fast/87416418657359cb625c412a48b6e1d6d41c29bd/chi_tra.traineddata'

$NeedsDownload = -not (Test-Path -LiteralPath $ChiTraPath)
if (-not $NeedsDownload) {
    $NeedsDownload = (Get-FileHash -Algorithm SHA256 -LiteralPath $ChiTraPath).Hash -ne $ChiTraHash
}
if ($NeedsDownload) {
    $TemporaryPath = "$ChiTraPath.download"
    try {
        Invoke-WebRequest -Uri $ChiTraUrl -OutFile $TemporaryPath
        $DownloadedHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $TemporaryPath).Hash
        if ($DownloadedHash -ne $ChiTraHash) {
            throw "Traditional Chinese OCR data hash mismatch: $DownloadedHash"
        }
        Move-Item -LiteralPath $TemporaryPath -Destination $ChiTraPath -Force
    }
    finally {
        if (Test-Path -LiteralPath $TemporaryPath) {
            Remove-Item -LiteralPath $TemporaryPath -Force
        }
    }
}

$env:TESSDATA_PREFIX = $TessData
# The 'User' scope only exists on Windows; .NET throws PlatformNotSupportedException elsewhere.
[Environment]::SetEnvironmentVariable('TESSDATA_PREFIX', $TessData, 'User')

Write-Host '[3/3] Verifying the Traditional Chinese model...'
$Languages = & $TesseractPath --list-langs 2>&1
if ($Languages -notcontains 'chi_tra') {
    throw "chi_tra is still missing from: $TessData"
}

Write-Host "Tesseract data: $TessData"
Write-Host 'OCR is ready. Use lang chi_tra+eng for Traditional Chinese handouts.'
