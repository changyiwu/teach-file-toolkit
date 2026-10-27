[CmdletBinding()]
param()

# Find or create a Python environment with the 12 core packages.
# The last line printed is the interpreter path.
# Requires PowerShell 7 (pwsh) on both Windows and macOS.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$env:PYTHONUTF8 = '1'

if ($null -eq $IsWindows) {
    throw 'PowerShell 7 (pwsh) is required; 5.1 has no $IsWindows and would silently take the wrong branch.'
}

$ScriptRoot = Split-Path -Parent $PSCommandPath
$RequirementsPath = Join-Path $ScriptRoot 'requirements-core.txt'

if ($IsWindows) {
    $SharedVenv = Join-Path $env:LOCALAPPDATA 'file-toolkit' '.venv'   # platform-ok: Windows 分支內
    $SharedPython = Join-Path $SharedVenv 'Scripts' 'python.exe'
    $LocalPython = Join-Path (Get-Location).Path '.venv' 'Scripts' 'python.exe'
} else {
    # ~/.local/share rather than ~/Library/Application Support: no spaces in the path,
    # which keeps quoting simple for every tool that receives this interpreter path.
    $SharedVenv = Join-Path $HOME '.local' 'share' 'file-toolkit' '.venv'
    $SharedPython = Join-Path $SharedVenv 'bin' 'python'
    $LocalPython = Join-Path (Get-Location).Path '.venv' 'bin' 'python'
}

$ImportCheck = 'import docx, openpyxl, pptx, pypdf, reportlab, PIL, matplotlib, qrcode, markitdown, docx2pdf, ocrmypdf, pypdfium2'

if (-not (Test-Path -LiteralPath $RequirementsPath)) {
    throw "Missing requirements file: $RequirementsPath"
}

function Test-PythonRuns {
    param([string]$PythonPath)

    if (-not $PythonPath -or -not (Test-Path -LiteralPath $PythonPath)) {
        return $false
    }

    try {
        & $PythonPath -c 'pass' 2>$null | Out-Null
        return ($LASTEXITCODE -eq 0)
    }
    catch {
        # Smart App Control blocks the launcher outright, so PowerShell throws instead of
        # returning an exit code. Treat that the same as "this interpreter cannot run".
        return $false
    }
}

function Test-CoreEnv {
    param([string]$PythonPath)

    if (-not (Test-PythonRuns -PythonPath $PythonPath)) {
        return $false
    }
    & $PythonPath -c $ImportCheck 2>$null | Out-Null
    return ($LASTEXITCODE -eq 0)
}

function Repair-VenvLauncher {
    # Windows only. Smart App Control blocks executables that have no established reputation,
    # and the small launcher uv writes into Scripts\ is one of them: the packages are fine, only
    # the entry point is unusable. Rebuilding with the base Python's own venv module drops a
    # signed python.exe in its place and leaves Lib/site-packages untouched, so the core
    # packages are kept rather than reinstalled.
    param([string]$VenvPath, [string]$PythonPath)

    $ConfigPath = Join-Path $VenvPath 'pyvenv.cfg'
    if (-not (Test-Path -LiteralPath $ConfigPath)) {
        return $false
    }

    # 'home' is the base Python folder; 'executable' is the interpreter itself. uv writes only
    # 'home', and the 'executable' path it would write can point at a version-pinned folder name
    # that no longer exists, so try 'home' first and require both existence and a working run.
    $HomeCandidates = @()
    $ExecutableCandidates = @()
    foreach ($line in (Get-Content -LiteralPath $ConfigPath)) {
        if ($line -match '^\s*home\s*=\s*(.+?)\s*$') {
            $HomeCandidates += (Join-Path $Matches[1] 'python.exe')   # platform-ok: Windows 專屬修復
        }
        elseif ($line -match '^\s*executable\s*=\s*(.+?)\s*$') {
            $ExecutableCandidates += $Matches[1]
        }
    }

    foreach ($candidate in ($HomeCandidates + $ExecutableCandidates)) {
        if (-not (Test-PythonRuns -PythonPath $candidate)) {
            continue
        }
        Write-Host 'The environment launcher cannot run (Smart App Control). Rebuilding it with the base Python...'
        & $candidate -m venv $VenvPath
        if ($LASTEXITCODE -eq 0 -and (Test-PythonRuns -PythonPath $PythonPath)) {
            Write-Host 'Launcher rebuilt; the installed packages were kept.'
            return $true
        }
    }

    return $false
}

function Find-Uv {
    $command = Get-Command 'uv' -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }

    if ($IsWindows) {
        $candidates = @(
            (Join-Path $env:USERPROFILE '.local' 'bin' 'uv.exe'),                  # platform-ok: Windows 分支內
            (Join-Path $env:LOCALAPPDATA 'Microsoft' 'WinGet' 'Links' 'uv.exe')    # platform-ok: Windows 分支內
        )
    } else {
        $candidates = @(
            (Join-Path $HOME '.local' 'bin' 'uv'),
            '/opt/homebrew/bin/uv',      # Apple Silicon Homebrew
            '/usr/local/bin/uv'          # Intel Homebrew
        )
    }
    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) { return $candidate }
    }
    return $null
}

# A shared environment whose launcher is blocked still holds every core package. Repair the
# entry point before the reuse check, or a healthy environment would be judged broken and
# rebuilt from scratch.
if ($IsWindows -and (Test-Path -LiteralPath $SharedPython) -and -not (Test-PythonRuns -PythonPath $SharedPython)) {
    Repair-VenvLauncher -VenvPath $SharedVenv -PythonPath $SharedPython | Out-Null
}

# Step 1: reuse an existing environment if one already has the core packages.
$Candidates = @(
    $env:FILE_TOOLKIT_PYTHON,
    $LocalPython,
    $SharedPython
)

foreach ($candidate in $Candidates) {
    if (Test-CoreEnv -PythonPath $candidate) {
        Write-Host 'Reusing existing core environment.'
        Write-Output (Resolve-Path -LiteralPath $candidate).Path
        exit 0
    }
}

# Step 2: no usable environment, so build the shared one.
Write-Host 'No usable core environment found. Creating the shared one...'
$UvPath = Find-Uv
if (-not $UvPath) {
    if ($IsWindows) {
        $WinGet = Get-Command 'winget' -ErrorAction SilentlyContinue
        if (-not $WinGet) {
            throw 'uv and WinGet were not found. Install uv from https://docs.astral.sh/uv/ and run this script again.'
        }
        Write-Host 'Installing astral-sh.uv with WinGet...'
        & $WinGet.Source install --exact --id astral-sh.uv --silent --accept-package-agreements --accept-source-agreements
    } else {
        $Brew = Get-Command 'brew' -ErrorAction SilentlyContinue
        if (-not $Brew) {
            throw 'uv and Homebrew were not found. Install uv from https://docs.astral.sh/uv/ and run this script again.'
        }
        Write-Host 'Installing uv with Homebrew...'
        & $Brew.Source install uv
    }
    $UvPath = Find-Uv
    if (-not $UvPath) {
        throw 'uv was installed, but this terminal cannot see the new PATH. Restart the Agent and run this script again.'
    }
}

# Windows on ARM: use x64 Python under Prism emulation. Natively, cryptography (pulled in by
# pdfminer.six), pi-heif and uharfbuzz have no win_arm64 wheels and fail to build.
# Check the OS architecture, not the process: x64 pwsh under emulation reports X64.
$IsWindowsArm = $IsWindows -and
    [Runtime.InteropServices.RuntimeInformation]::OSArchitecture -eq [Runtime.InteropServices.Architecture]::Arm64
$PythonRequest = if ($IsWindowsArm) { 'cpython-3.12-windows-x86_64-none' } else { '3.12' }

if (-not (Test-Path -LiteralPath $SharedPython)) {
    # Keep uv's own output: on a fresh machine this step also downloads Python, and a bare
    # "failed" message leaves nothing to diagnose once a retry succeeds.
    $UvOutput = & $UvPath venv --python $PythonRequest $SharedVenv --quiet 2>&1 | Out-String
    $UvExit = $LASTEXITCODE
    if ($UvExit -ne 0 -or -not (Test-Path -LiteralPath $SharedPython)) {
        throw "Failed to create the shared environment at $SharedVenv (uv exit code $UvExit, interpreter exists: $(Test-Path -LiteralPath $SharedPython)). uv output: $($UvOutput.Trim())"
    }
}

# uv pip install runs this interpreter, so the launcher has to work before the next step.
if ($IsWindows -and -not (Test-PythonRuns -PythonPath $SharedPython)) {
    if (-not (Repair-VenvLauncher -VenvPath $SharedVenv -PythonPath $SharedPython)) {
        throw "The interpreter at $SharedPython cannot run, most likely blocked by Smart App Control, and rebuilding its launcher failed. Delete $SharedVenv and run this script again."
    }
}

Write-Host 'Installing 12 core packages (media tools and advanced Office automation are excluded)...'
& $UvPath pip install --python $SharedPython --requirements $RequirementsPath --quiet --no-progress
if ($LASTEXITCODE -ne 0) {
    throw 'Core package installation failed. Keep the original error above; do not install every optional package as a workaround.'
}

if (-not (Test-CoreEnv -PythonPath $SharedPython)) {
    throw 'Core packages were installed, but the import check still fails.'
}

Write-Host 'Shared core environment is ready.'
Write-Output (Resolve-Path -LiteralPath $SharedPython).Path
