[CmdletBinding()]
param(
    [switch]$SkipTests
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$ProjectRoot = (Resolve-Path -LiteralPath $PSScriptRoot).Path
$Python = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
$BuildRoot = [IO.Path]::GetFullPath((Join-Path $ProjectRoot "build\pyinstaller"))
$DistRoot = [IO.Path]::GetFullPath((Join-Path $ProjectRoot "dist"))
$FinalDist = [IO.Path]::GetFullPath((Join-Path $DistRoot "TVC_JOB_BOT_v0.8.0"))
$WorkerDistRoot = Join-Path $BuildRoot "worker-dist"
$ControlDistRoot = Join-Path $BuildRoot "control-dist"
$WorkerFolder = Join-Path $WorkerDistRoot "TVC Bot Worker"
$ControlFolder = Join-Path $ControlDistRoot "TVC Bot Control"

function Assert-SafeProjectPath {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$ExpectedRelativePath
    )

    $fullPath = [IO.Path]::GetFullPath($Path)
    $expectedPath = [IO.Path]::GetFullPath((Join-Path $ProjectRoot $ExpectedRelativePath))
    if (-not $fullPath.Equals($expectedPath, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing unsafe build path: $fullPath"
    }
    if (-not $fullPath.StartsWith($ProjectRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Build path escaped project root: $fullPath"
    }
    if (Test-Path -LiteralPath $fullPath) {
        $item = Get-Item -LiteralPath $fullPath -Force
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "Refusing to clean reparse-point path: $fullPath"
        }
    }
    return $fullPath
}

function Invoke-PyInstallerSpec {
    param(
        [Parameter(Mandatory = $true)][string]$Spec,
        [Parameter(Mandatory = $true)][string]$WorkPath,
        [Parameter(Mandatory = $true)][string]$DistPath
    )

    & $Python -m PyInstaller `
        --noconfirm `
        --clean `
        --workpath $WorkPath `
        --distpath $DistPath `
        $Spec
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller failed for $Spec with exit code $LASTEXITCODE"
    }
}

function Merge-DirectoryVerified {
    param(
        [Parameter(Mandatory = $true)][string]$Source,
        [Parameter(Mandatory = $true)][string]$Destination
    )

    if (-not (Test-Path -LiteralPath $Source -PathType Container)) {
        throw "Missing build output directory: $Source"
    }
    $sourceRoot = [IO.Path]::GetFullPath($Source).TrimEnd(
        [IO.Path]::DirectorySeparatorChar,
        [IO.Path]::AltDirectorySeparatorChar
    ) + [IO.Path]::DirectorySeparatorChar
    New-Item -ItemType Directory -Force -Path $Destination | Out-Null
    foreach ($file in Get-ChildItem -LiteralPath $Source -Recurse -File) {
        $filePath = [IO.Path]::GetFullPath($file.FullName)
        if (-not $filePath.StartsWith($sourceRoot, [StringComparison]::OrdinalIgnoreCase)) {
            throw "Build output escaped source folder: $filePath"
        }
        $relative = $filePath.Substring($sourceRoot.Length)
        $target = Join-Path $Destination $relative
        $targetParent = Split-Path -Parent $target
        New-Item -ItemType Directory -Force -Path $targetParent | Out-Null
        if (Test-Path -LiteralPath $target -PathType Leaf) {
            $sourceHash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
            $targetHash = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash
            if ($sourceHash -ne $targetHash) {
                $normalizedRelative = $relative.Replace("/", "\")
                if ($normalizedRelative -ne "_internal\base_library.zip") {
                    throw "Dependency collision differs: $relative"
                }
                & $Python -c @"
import hashlib
import sys
import zipfile

def content(path):
    with zipfile.ZipFile(path) as archive:
        return {
            item.filename: hashlib.sha256(archive.read(item)).digest()
            for item in archive.infolist()
            if not item.is_dir()
        }

raise SystemExit(0 if content(sys.argv[1]) == content(sys.argv[2]) else 1)
"@ $file.FullName $target
                if ($LASTEXITCODE -ne 0) {
                    throw "base_library.zip contents differ: $relative"
                }
            }
            continue
        }
        Copy-Item -LiteralPath $file.FullName -Destination $target
    }
}

if (-not (Test-Path -LiteralPath $Python -PathType Leaf)) {
    throw "Project Python not found: $Python"
}

& $Python -c "import PyInstaller; print('PyInstaller', PyInstaller.__version__)"
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller is missing. Run: .venv\Scripts\python.exe -m pip install -r requirements-build.txt"
}

if (-not $SkipTests) {
    & $Python -m compileall -q (Join-Path $ProjectRoot "src") (Join-Path $ProjectRoot "tests")
    if ($LASTEXITCODE -ne 0) { throw "Python compile check failed" }
    & $Python -m unittest discover -s (Join-Path $ProjectRoot "tests") -p "test_*.py"
    if ($LASTEXITCODE -ne 0) { throw "Unit/mock suite failed" }
}

$BuildRoot = Assert-SafeProjectPath $BuildRoot "build\pyinstaller"
$FinalDist = Assert-SafeProjectPath $FinalDist "dist\TVC_JOB_BOT_v0.8.0"

if (Test-Path -LiteralPath $BuildRoot) {
    Remove-Item -LiteralPath $BuildRoot -Recurse -Force
}
if (Test-Path -LiteralPath $FinalDist) {
    Remove-Item -LiteralPath $FinalDist -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $BuildRoot, $DistRoot, $FinalDist | Out-Null

Write-Host "Building TVC Bot Worker.exe..."
Invoke-PyInstallerSpec `
    (Join-Path $ProjectRoot "build_specs\tvc_bot_worker.spec") `
    (Join-Path $BuildRoot "worker-work") `
    $WorkerDistRoot

Write-Host "Building TVC Bot Control.exe..."
Invoke-PyInstallerSpec `
    (Join-Path $ProjectRoot "build_specs\tvc_bot_control.spec") `
    (Join-Path $BuildRoot "control-work") `
    $ControlDistRoot

# Each spec is a valid standalone ONEDIR build. Merge them only after checking
# that overlapping dependency files are byte-identical.
Merge-DirectoryVerified $ControlFolder $FinalDist
Merge-DirectoryVerified $WorkerFolder $FinalDist

Copy-Item -LiteralPath (Join-Path $ProjectRoot "config.ini") -Destination $FinalDist -Force
Copy-Item -LiteralPath (Join-Path $ProjectRoot "field_map.json") -Destination $FinalDist -Force
Copy-Item -LiteralPath (Join-Path $ProjectRoot "assets") -Destination $FinalDist -Recurse -Force

$RequiredFiles = @(
    (Join-Path $FinalDist "TVC Bot Control.exe"),
    (Join-Path $FinalDist "TVC Bot Worker.exe"),
    (Join-Path $FinalDist "config.ini"),
    (Join-Path $FinalDist "field_map.json")
)
foreach ($required in $RequiredFiles) {
    if (-not (Test-Path -LiteralPath $required -PathType Leaf)) {
        throw "Final distribution is missing: $required"
    }
}
if (-not (Test-Path -LiteralPath (Join-Path $FinalDist "_internal") -PathType Container)) {
    throw "Final distribution is missing the _internal directory"
}
if (-not (Test-Path -LiteralPath (Join-Path $FinalDist "assets") -PathType Container)) {
    throw "Final distribution is missing the assets directory"
}

$OptionalAssets = @(
    "bot_ready.png",
    "bot_running.png",
    "bot_success.png",
    "bot_error.png",
    "app_icon.ico"
)
foreach ($asset in $OptionalAssets) {
    $assetPath = Join-Path $FinalDist "assets\$asset"
    if (-not (Test-Path -LiteralPath $assetPath -PathType Leaf)) {
        Write-Warning "Optional GUI asset is missing; runtime fallback will be used: $asset"
    }
}

Write-Host "Running mandatory frozen Worker smoke tests..."
& $Python `
    (Join-Path $ProjectRoot "tests\frozen_worker_smoke.py") `
    --worker (Join-Path $FinalDist "TVC Bot Worker.exe")
if ($LASTEXITCODE -ne 0) {
    throw "Frozen Worker post-build smoke failed with exit code $LASTEXITCODE"
}

Write-Host ""
Write-Host "Build complete: $FinalDist"
Get-Item -LiteralPath `
    (Join-Path $FinalDist "TVC Bot Control.exe"), `
    (Join-Path $FinalDist "TVC Bot Worker.exe") |
    Select-Object Name, Length, FullName |
    Format-Table -AutoSize
