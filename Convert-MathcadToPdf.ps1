
param(
    [Parameter(Mandatory = $true)] [string] $InputDir,
    [Parameter(Mandatory = $true)] [string] $OutputDir
)

$ErrorActionPreference = 'Stop'

# Validate dirs
if (-not (Test-Path $InputDir)) { Write-Error "Input directory not found: $InputDir"; exit 1 }
if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir | Out-Null }

function Join-PathSafe {
    param([Parameter(Mandatory = $true)][string]$a, [Parameter(Mandatory = $true)][string]$b)
    Join-Path -Path $a -ChildPath $b
}

# Enumerate .mcdx
$files = Get-ChildItem -Path $InputDir -Filter *.mcdx -File -Recurse
if (-not $files) { Write-Host "No .mcdx files found in '$InputDir'. Nothing to do."; exit 0 }

# ---- Mathcad Prime 11 paths ----
$mcRoot = "C:\Program Files\PTC\Mathcad Prime 11.0.0.0"
$mcDll  = Join-Path $mcRoot "Ptc.MathcadPrime.Automation.dll"
if (-not (Test-Path $mcDll)) { Write-Error "Automation DLL not found: $mcDll"; exit 1 }

# Load assembly and get coclass
$asm = [Reflection.Assembly]::LoadFrom($mcDll)
$creatorType = $asm.GetType('Ptc.MathcadPrime.Automation.ApplicationCreatorClass', $true)
if (-not $creatorType) { Write-Error "ApplicationCreatorClass not found in the Automation assembly."; exit 1 }

# Create app
$primeApp = [System.Activator]::CreateInstance($creatorType)

foreach ($f in $files) {
    $ws = $null
    $outPdf = Join-PathSafe $OutputDir ($f.BaseName + '.pdf')
    Write-Host "Converting: $($f.FullName) -> $outPdf"

    try {
        # Open worksheet (IMathcadPrimeWorksheet3)
        $ws = $primeApp.Open($f.FullName)
        if (-not $ws) { throw "Open() returned null for: $($f.FullName)" }

        # Optional: ensure calculation; API has calc/timeouts depending on version
        # $ws.DefaultCalculationTimeout(60)

        # Save directly to PDF (Prime 6.0+ supports via extension)
        $ws.SaveAs($outPdf)
        Write-Host "Saved: $outPdf"
    }
    catch {
        Write-Warning ("Failed to convert '{0}': {1}" -f $f.FullName, $_.Exception.Message)
    }
    finally {
        if ($ws) {
            try { $ws.Close([Ptc.MathcadPrime.Automation.SaveOption]::spDiscardChanges) } catch {}
            $ws = $null
        }
    }
}