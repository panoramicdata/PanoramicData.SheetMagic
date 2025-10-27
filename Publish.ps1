#Requires -Version 7.0

<#
.SYNOPSIS
    Publishes PanoramicData.SheetMagic NuGet package after validation.

.DESCRIPTION
    This script performs the following steps:
    1. Checks Git working directory is clean (no uncommitted changes)
    2. Runs all unit tests
    3. Builds the package
    4. Publishes to NuGet.org using API key from nuget-key.txt

.PARAMETER SkipTests
    Skip running unit tests (not recommended for production use)

.PARAMETER DryRun
    Perform all checks but don't actually publish to NuGet

.EXAMPLE
    .\Publish.ps1
    
.EXAMPLE
    .\Publish.ps1 -DryRun
#>

[CmdletBinding()]
param(
    [switch]$SkipTests,
  [switch]$DryRun
)

$ErrorActionPreference = "Stop"
$InformationPreference = "Continue"

# Colors for output
function Write-Success { param($Message) Write-Host "? $Message" -ForegroundColor Green }
function Write-Error { param($Message) Write-Host "? $Message" -ForegroundColor Red }
function Write-Step { param($Message) Write-Host "`n==> $Message" -ForegroundColor Cyan }

# Script variables
$ScriptRoot = $PSScriptRoot
$ProjectPath = Join-Path $ScriptRoot "PanoramicData.SheetMagic\PanoramicData.SheetMagic.csproj"
$TestProjectPath = Join-Path $ScriptRoot "PanoramicData.SheetMagic.Test\PanoramicData.SheetMagic.Test.csproj"
$NuGetKeyFile = Join-Path $ScriptRoot "nuget-key.txt"
$NuGetSource = "https://api.nuget.org/v3/index.json"

# Ensure we're in the correct directory
Set-Location $ScriptRoot

try {
    Write-Host "`n????????????????????????????????????????????????????????????" -ForegroundColor Cyan
    Write-Host "?  PanoramicData.SheetMagic - NuGet Publish Script     ?" -ForegroundColor Cyan
 Write-Host "????????????????????????????????????????????????????????????`n" -ForegroundColor Cyan

    # Step 1: Check Git status
    Write-Step "Checking Git working directory status..."
    
    $gitStatus = git status --porcelain 2>&1
    if ($LASTEXITCODE -ne 0) {
     throw "Failed to check Git status. Is this a Git repository?"
    }
    
    if ($gitStatus) {
  Write-Error "Git working directory is not clean. Please commit or stash changes first."
        Write-Host "`nUncommitted changes:" -ForegroundColor Yellow
     Write-Host $gitStatus
        exit 1
    }
    
    Write-Success "Git working directory is clean"
    
 # Get current branch and latest commit
    $currentBranch = git rev-parse --abbrev-ref HEAD
    $latestCommit = git rev-parse --short HEAD
    Write-Information "Branch: $currentBranch | Commit: $latestCommit"

    # Step 2: Run unit tests (unless skipped)
    if (-not $SkipTests) {
Write-Step "Running unit tests..."
        
        dotnet test $TestProjectPath --configuration Release --nologo
        
        if ($LASTEXITCODE -ne 0) {
        Write-Error "Unit tests failed. Fix the tests before publishing."
      exit 1
        }
   
        Write-Success "All unit tests passed"
    }
    else {
        Write-Host "? Skipping unit tests (not recommended)" -ForegroundColor Yellow
    }

    # Step 3: Build and pack the project
    Write-Step "Building and packing the project..."
    
    dotnet pack $ProjectPath --configuration Release --nologo
    
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Build/pack failed"
        exit 1
    }
    
    Write-Success "Project built and packed successfully"

    # Step 4: Find the generated .nupkg file
    Write-Step "Locating NuGet package..."
  
    $packagePath = Get-ChildItem -Path (Join-Path $ScriptRoot "PanoramicData.SheetMagic\bin\Release") -Filter "*.nupkg" -Recurse | 
        Where-Object { $_.Name -notlike "*.symbols.nupkg" } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if (-not $packagePath) {
        Write-Error "Could not find generated NuGet package"
        exit 1
    }

    $packageName = $packagePath.Name
    Write-Success "Found package: $packageName"
 Write-Information "Package path: $($packagePath.FullName)"

    # Step 5: Check for NuGet API key
    Write-Step "Checking NuGet API key..."
    
 if (-not (Test-Path $NuGetKeyFile)) {
        Write-Error "NuGet API key file not found: $NuGetKeyFile"
  Write-Host "`nPlease create the file and add your NuGet API key to it." -ForegroundColor Yellow
        Write-Host "You can get an API key from: https://www.nuget.org/account/apikeys" -ForegroundColor Yellow
        exit 1
    }

    $apiKey = Get-Content $NuGetKeyFile -Raw | ForEach-Object { $_.Trim() }
    
  if ([string]::IsNullOrWhiteSpace($apiKey)) {
   Write-Error "NuGet API key file is empty: $NuGetKeyFile"
        Write-Host "`nPlease add your NuGet API key to the file." -ForegroundColor Yellow
        Write-Host "You can get an API key from: https://www.nuget.org/account/apikeys" -ForegroundColor Yellow
        exit 1
    }

    Write-Success "NuGet API key loaded"

    # Step 6: Publish to NuGet (or dry run)
    if ($DryRun) {
        Write-Host "`n? DRY RUN MODE - Package will NOT be published" -ForegroundColor Yellow
  Write-Host "`nWould publish:" -ForegroundColor Cyan
    Write-Host "  Package: $packageName" -ForegroundColor White
        Write-Host "  Source:  $NuGetSource" -ForegroundColor White
        Write-Success "Dry run completed successfully"
    }
    else {
        Write-Step "Publishing to NuGet.org..."
        
        # Confirm before publishing
        Write-Host "`nAbout to publish:" -ForegroundColor Yellow
        Write-Host "  Package: $packageName" -ForegroundColor White
        Write-Host "  Source:  $NuGetSource" -ForegroundColor White
        
      $confirmation = Read-Host "`nDo you want to continue? (yes/no)"
        
        if ($confirmation -ne "yes") {
            Write-Host "`n? Publish cancelled by user" -ForegroundColor Yellow
            exit 0
        }

  dotnet nuget push $packagePath.FullName --api-key $apiKey --source $NuGetSource --skip-duplicate
        
   if ($LASTEXITCODE -ne 0) {
            Write-Error "Failed to publish package to NuGet"
      exit 1
        }

        Write-Success "Package published successfully to NuGet.org!"
        Write-Host "`n?? Package $packageName is now available on NuGet.org" -ForegroundColor Green
        Write-Host "   It may take a few minutes to appear in search results.`n" -ForegroundColor Gray
    }

}
catch {
    Write-Host "`n? Error: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    exit 1
}
finally {
    # Clean sensitive data from memory
    if ($apiKey) {
        Clear-Variable apiKey -ErrorAction SilentlyContinue
    }
}
