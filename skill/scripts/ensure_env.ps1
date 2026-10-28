[CmdletBinding()]
param(
    # Re-install the pinned list into a reused environment too (install_windows.ps1 uses this).
    [switch]$Sync,
    # Windows on ARM only: try native ARM64 again for an existing x64 shared environment. The x64
    # one is kept aside and restored if native still fails, so this is safe to run any time.
    [switch]$RetryNative
)

# Find or create a Python environment with the 12 core packages.
# The last line printed is the interpreter path.
# Requires PowerShell 7 (pwsh) on both Windows and macOS.
#
# Three platforms, one shared environment path each:
#   Windows x64        -> x64 Python
#   Windows ARM64      -> native ARM64 Python first; if it cannot install or Smart App Control
#                         blocks its native modules, rebuild as x64 Python under Prism emulation
#                         (-RetryNative tries native again later, once SAC reputation has grown)
#   macOS Apple Silicon -> native arm64 Python (Intel Macs get whatever uv picks for 3.12)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$env:PYTHONUTF8 = '1'

if ($null -eq $IsWindows) {
    throw 'PowerShell 7 (pwsh) is required; 5.1 has no $IsWindows and would silently take the wrong branch.'
}

$ScriptRoot = Split-Path -Parent $PSCommandPath
$RequirementsPath = Join-Path $ScriptRoot 'requirements-core.txt'
$ArmOverridesPath = Join-Path $ScriptRoot 'overrides-win-arm64.txt'

# Same place on both platforms. Not %LOCALAPPDATA% on Windows: the Claude desktop app is an MSIX
# package, and AppData writes made from inside it land in its private
# Packages\Claude_*\LocalCache folder, invisible to Explorer, terminals and other agents
# (2026-10-28, LT-YI-RS). Not ~/Library/Application Support on macOS: no spaces in the path,
# which keeps quoting simple for every tool that receives this interpreter path.
$SharedVenv = Join-Path $HOME '.local' 'share' 'file-toolkit' '.venv'
if ($IsWindows) {
    $SharedPython = Join-Path $SharedVenv 'Scripts' 'python.exe'
    $LocalPython = Join-Path (Get-Location).Path '.venv' 'Scripts' 'python.exe'
    # The venv points at its base Python by absolute path, and uv's default for that is
    # %APPDATA%\uv\python, which the same redirection would hide. Keep it next to the venv.
    if (-not $env:UV_PYTHON_INSTALL_DIR) {
        $env:UV_PYTHON_INSTALL_DIR = Join-Path $HOME '.local' 'share' 'uv' 'python'
    }
} else {
    $SharedPython = Join-Path $SharedVenv 'bin' 'python'
    $LocalPython = Join-Path (Get-Location).Path '.venv' 'bin' 'python'
}

# Importing is not enough: Smart App Control blocks individual native modules (.pyd) that load
# lazily, e.g. matplotlib's Agg backend on a native ARM64 build (2026-10-28, LT-YI-RS). So actually
# draw a chart, contour, save PDF, render a PDF page and load COM, the way the recipes do.
# Single quotes only, so the code survives native argument passing on every pwsh 7 version.
$CoreCheck = @'
import io, sys
import docx, openpyxl, pptx, pypdf, reportlab, PIL, matplotlib, qrcode, markitdown, docx2pdf, ocrmypdf, pypdfium2
import numpy, pandas, lxml.etree, pikepdf, cryptography.hazmat.primitives.ciphers
matplotlib.use('Agg')
from matplotlib import pyplot as plt
fig, ax = plt.subplots()
ax.plot([0, 1], [0, 1])
ax.contour([[0, 1], [1, 0]])
fig.savefig(io.BytesIO(), format='png')
fig.savefig(io.BytesIO(), format='pdf')
pdf = pypdfium2.PdfDocument.new()
pdf.new_page(72, 72)
pdf[0].render().to_pil().save(io.BytesIO(), 'PNG')
if sys.platform == 'win32':
    import win32com.client, pythoncom
'@

foreach ($path in @($RequirementsPath, $ArmOverridesPath)) {
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Missing file: $path"
    }
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

function Get-CoreEnvError {
    # $null when the environment passes $CoreCheck, otherwise the last line of the error.
    param([string]$PythonPath)

    if (-not (Test-PythonRuns -PythonPath $PythonPath)) {
        return "the interpreter cannot run: $PythonPath"
    }
    $Output = & $PythonPath -c $CoreCheck 2>&1 | Out-String
    if ($LASTEXITCODE -eq 0) { return $null }
    $Lines = @($Output.Trim() -split "`r?`n")
    return $Lines[-1]
}

function Test-CoreEnv {
    param([string]$PythonPath)
    return ($null -eq (Get-CoreEnvError -PythonPath $PythonPath))
}

function Get-PythonPlatform {
    # 'win-amd64', 'win-arm64', 'macosx-...-arm64'; $null when the interpreter cannot run.
    # Not platform.machine(): x64 Python under Prism reports the host's ARM64 there.
    param([string]$PythonPath)

    if (-not (Test-PythonRuns -PythonPath $PythonPath)) { return $null }
    return (& $PythonPath -c 'import sysconfig; print(sysconfig.get_platform())' 2>$null | Out-String).Trim()
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
        # --without-pip: uv environments have no pip, and ensurepip itself can fail here
        # (2026-10-28, LT-YI-RS x64 Python), which would sink an otherwise good repair.
        & $candidate -m venv --without-pip $VenvPath
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

# Extra uv pip arguments for an interpreter: native Windows ARM64 needs the overrides file that
# drops uharfbuzz (no win_arm64 wheel); everything else installs the list as is.
function Get-InstallArgs {
    param([string]$PythonPath)
    if ((Get-PythonPlatform -PythonPath $PythonPath) -eq 'win-arm64') {
        return @('--overrides', $ArmOverridesPath)
    }
    return @()
}

function Install-CorePackages {
    # $null on success, otherwise why it failed.
    param([string]$UvPath, [string]$PythonPath)

    $InstallArgs = @(Get-InstallArgs -PythonPath $PythonPath)
    $UvOutput = & $UvPath pip install --python $PythonPath --requirements $RequirementsPath @InstallArgs --quiet --no-progress 2>&1 | Out-String
    if ($LASTEXITCODE -ne 0) {
        return "uv pip install failed (exit code $LASTEXITCODE): $($UvOutput.Trim())"
    }
    return (Get-CoreEnvError -PythonPath $PythonPath)
}

function New-SharedEnv {
    # Create (if missing) and fill the shared environment with the given uv Python request.
    # $null on success, otherwise why it failed, so the ARM64 attempt can fall back to x64.
    param([string]$UvPath, [string]$PythonRequest)

    if (-not (Test-Path -LiteralPath $SharedPython)) {
        # Keep uv's own output: on a fresh machine this step also downloads Python, and a bare
        # "failed" message leaves nothing to diagnose once a retry succeeds.
        $UvOutput = & $UvPath venv --python $PythonRequest $SharedVenv --quiet 2>&1 | Out-String
        $UvExit = $LASTEXITCODE
        if ($UvExit -ne 0 -or -not (Test-Path -LiteralPath $SharedPython)) {
            return "Failed to create the shared environment at $SharedVenv (uv exit code $UvExit, interpreter exists: $(Test-Path -LiteralPath $SharedPython)). uv output: $($UvOutput.Trim())"
        }
    }

    # uv pip install runs this interpreter, so the launcher has to work before the next step.
    if ($IsWindows -and -not (Test-PythonRuns -PythonPath $SharedPython)) {
        if (-not (Repair-VenvLauncher -VenvPath $SharedVenv -PythonPath $SharedPython)) {
            return "The interpreter at $SharedPython cannot run, most likely blocked by Smart App Control, and rebuilding its launcher failed. Delete $SharedVenv and run this script again."
        }
    }

    Write-Host 'Installing 12 core packages (media tools and advanced Office automation are excluded)...'
    return (Install-CorePackages -UvPath $UvPath -PythonPath $SharedPython)
}

# A shared environment whose launcher is blocked still holds every core package. Repair the
# entry point before the reuse check, or a healthy environment would be judged broken and
# rebuilt from scratch.
if ($IsWindows -and (Test-Path -LiteralPath $SharedPython) -and -not (Test-PythonRuns -PythonPath $SharedPython)) {
    Repair-VenvLauncher -VenvPath $SharedVenv -PythonPath $SharedPython | Out-Null
}

# Check the OS architecture, not the process: x64 pwsh under emulation reports X64 as its process.
$IsArmHost = [Runtime.InteropServices.RuntimeInformation]::OSArchitecture -eq [Runtime.InteropServices.Architecture]::Arm64
$X64Request = 'cpython-3.12-windows-x86_64-none'
$WinArmRequest = 'cpython-3.12-windows-aarch64-none'
# Renamed, not copied: console-script launchers hard-code the venv path, so a moved copy would
# break, while renaming back restores it exactly.
$X64Backup = "$SharedVenv-x64-backup"

# An interrupted -RetryNative leaves the x64 environment parked under the backup name. Keep the
# shared one if it passes; otherwise the backup is the last known good environment, so put it back.
if ($IsWindows -and (Test-Path -LiteralPath $X64Backup)) {
    if (Test-CoreEnv -PythonPath $SharedPython) {
        Remove-Item -LiteralPath $X64Backup -Recurse -Force
    }
    else {
        Write-Host 'Restoring the x64 environment left aside by an interrupted -RetryNative...'
        if (Test-Path -LiteralPath $SharedVenv) {
            Remove-Item -LiteralPath $SharedVenv -Recurse -Force
        }
        Rename-Item -LiteralPath $X64Backup -NewName (Split-Path -Leaf $SharedVenv)
    }
}

if ($RetryNative) {
    if (-not ($IsWindows -and $IsArmHost)) {
        Write-Host '-RetryNative only applies to Windows on ARM; ignored.'
    }
    elseif ((Get-PythonPlatform -PythonPath $SharedPython) -ne 'win-amd64') {
        # Already native, or no x64 environment to replace: the normal flow below tries native first.
        Write-Host '-RetryNative: the shared environment is not x64, nothing to retry.'
    }
    else {
        $UvPath = Find-Uv
        if (-not $UvPath) {
            throw '-RetryNative needs uv, which was not found. Run this script once without -RetryNative.'
        }
        Write-Host '-RetryNative: setting the x64 environment aside and trying native ARM64...'
        try {
            Rename-Item -LiteralPath $SharedVenv -NewName (Split-Path -Leaf $X64Backup)
        }
        catch {
            throw "Could not set $SharedVenv aside, most likely because a program is still using it. Close it and run again. $($_.Exception.Message)"
        }
        $Failure = New-SharedEnv -UvPath $UvPath -PythonRequest $WinArmRequest
        if ($Failure) {
            Write-Host "Native ARM64 is still not usable: $Failure"
            Write-Host 'Keeping the x64 environment.'
            if (Test-Path -LiteralPath $SharedVenv) {
                Remove-Item -LiteralPath $SharedVenv -Recurse -Force
            }
            Rename-Item -LiteralPath $X64Backup -NewName (Split-Path -Leaf $SharedVenv)
        }
        else {
            Write-Host 'Native ARM64 passed the core check; switched to it and removed the x64 environment.'
            Remove-Item -LiteralPath $X64Backup -Recurse -Force
        }
    }
}

# Step 1: reuse an existing environment if one already passes the core check, whatever its
# architecture. An x64 environment on an ARM machine stays x64 until -RetryNative succeeds.
$Candidates = @(
    $env:FILE_TOOLKIT_PYTHON,
    $LocalPython,
    $SharedPython
)

foreach ($candidate in $Candidates) {
    if (-not (Test-CoreEnv -PythonPath $candidate)) {
        continue
    }
    Write-Host "Reusing existing core environment ($(Get-PythonPlatform -PythonPath $candidate))."
    if ($Sync) {
        # Reuse passes with older pins too, so -Sync brings the versions back in line with the list.
        $UvPath = Find-Uv
        if ($UvPath) {
            Write-Host 'Syncing the 12 core packages with the pinned list...'
            $Failure = Install-CorePackages -UvPath $UvPath -PythonPath $candidate
            if ($Failure) {
                throw "Core package sync failed. Keep this error; do not install every optional package as a workaround. $Failure"
            }
        }
        else {
            Write-Host 'uv was not found, so the pinned versions were not re-synced. The existing environment passed the core check and is used as is.'
        }
    }
    Write-Output (Resolve-Path -LiteralPath $candidate).Path
    exit 0
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

if ($IsWindows -and $IsArmHost) {
    # A leftover x64 environment here only failed the check (missing or broken packages), so
    # repair it as x64 instead of throwing it away for another native attempt.
    if ((Get-PythonPlatform -PythonPath $SharedPython) -eq 'win-amd64') {
        Write-Host 'Windows ARM64: repairing the existing x64 environment...'
        $Failure = New-SharedEnv -UvPath $UvPath -PythonRequest $X64Request
    }
    else {
        Write-Host 'Windows ARM64: trying a native ARM64 environment first...'
        $Failure = New-SharedEnv -UvPath $UvPath -PythonRequest $WinArmRequest
        if ($Failure) {
            # Native ARM64 wheels are downloaded far less, so Smart App Control often has no
            # reputation for their DLLs yet (2026-10-28, LT-YI-RS: 28 of 152 modules blocked).
            Write-Host "Native ARM64 environment is not usable: $Failure"
            Write-Host 'Falling back to x64 Python under Prism emulation...'
            if (Test-Path -LiteralPath $SharedVenv) {
                Remove-Item -LiteralPath $SharedVenv -Recurse -Force
            }
            $Failure = New-SharedEnv -UvPath $UvPath -PythonRequest $X64Request
        }
    }
}
elseif ($IsMacOS -and $IsArmHost) {
    # Ask for arm64 explicitly so a Rosetta-launched shell does not end up with an x86_64 Python.
    $Failure = New-SharedEnv -UvPath $UvPath -PythonRequest 'cpython-3.12-macos-aarch64-none'
}
else {
    $Failure = New-SharedEnv -UvPath $UvPath -PythonRequest '3.12'
}

if ($Failure) {
    throw "Core environment setup failed. Keep this error; do not install every optional package as a workaround. $Failure"
}

Write-Host "Shared core environment is ready ($(Get-PythonPlatform -PythonPath $SharedPython))."
Write-Output (Resolve-Path -LiteralPath $SharedPython).Path
