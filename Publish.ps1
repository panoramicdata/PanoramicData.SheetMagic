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

# Helper functions for colored output using Write-Information
function Write-Success { 
    param($Message) 
    Write-Information "? $Message" -InformationAction Continue
    Write-Verbose "? $Message" -Verbose
}

function Write-ErrorMessage { 
    param($Message) 
    Write-Error "? $Message" -ErrorAction Continue
}

function Write-Step { 
    param($Message) 
    Write-Information "`n==> $Message" -InformationAction Continue
}

# Script variables
$ScriptRoot = $PSScriptRoot
$ProjectPath = Join-Path $ScriptRoot "PanoramicData.SheetMagic\PanoramicData.SheetMagic.csproj"
$TestProjectPath = Join-Path $ScriptRoot "PanoramicData.SheetMagic.Test\PanoramicData.SheetMagic.Test.csproj"
$NuGetKeyFile = Join-Path $ScriptRoot "nuget-key.txt"
$NuGetSource = "https://api.nuget.org/v3/index.json"

# Ensure we're in the correct directory
Set-Location $ScriptRoot

try {
    Write-Information "`n??????????????????????????????????????????????????????" -InformationAction Continue
    Write-Information "?  PanoramicData.SheetMagic - NuGet Publish Script     ?" -InformationAction Continue
    Write-Information "??????????????????????????????????????????????????????`n" -InformationAction Continue

    # Step 1: Check Git status
    Write-Step "Checking Git working directory status..."
    
    $gitStatus = git status --porcelain 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to check Git status. Is this a Git repository?"
    }
    
    if ($gitStatus) {
        Write-ErrorMessage "Git working directory is not clean. Please commit or stash changes first."
        Write-Warning "`nUncommitted changes:"
        Write-Output $gitStatus
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
            Write-ErrorMessage "Unit tests failed. Fix the tests before publishing."
            exit 1
        }
   
        Write-Success "All unit tests passed"
    }
    else {
        Write-Warning "? Skipping unit tests (not recommended)"
    }

    # Step 3: Build and pack the project
    Write-Step "Building and packing the project..."
    
    dotnet pack $ProjectPath --configuration Release --nologo
    
    if ($LASTEXITCODE -ne 0) {
        Write-ErrorMessage "Build/pack failed"
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
        Write-ErrorMessage "Could not find generated NuGet package"
        exit 1
    }

    $packageName = $packagePath.Name
    Write-Success "Found package: $packageName"
    Write-Information "Package path: $($packagePath.FullName)"

    # Step 5: Check for NuGet API key
    Write-Step "Checking NuGet API key..."
    
    if (-not (Test-Path $NuGetKeyFile)) {
        Write-ErrorMessage "NuGet API key file not found: $NuGetKeyFile"
        Write-Warning "`nPlease create the file and add your NuGet API key to it."
        Write-Information "You can get an API key from: https://www.nuget.org/account/apikeys"
        exit 1
    }

    $apiKey = Get-Content $NuGetKeyFile -Raw | ForEach-Object { $_.Trim() }
    
    if ([string]::IsNullOrWhiteSpace($apiKey)) {
        Write-ErrorMessage "NuGet API key file is empty: $NuGetKeyFile"
        Write-Warning "`nPlease add your NuGet API key to the file."
        Write-Information "You can get an API key from: https://www.nuget.org/account/apikeys"
        exit 1
    }

    Write-Success "NuGet API key loaded"

    # Step 6: Publish to NuGet (or dry run)
    if ($DryRun) {
        Write-Warning "`n? DRY RUN MODE - Package will NOT be published"
        Write-Information "`nWould publish:" -InformationAction Continue
        Write-Information "  Package: $packageName" -InformationAction Continue
        Write-Information "  Source:  $NuGetSource" -InformationAction Continue
        Write-Success "Dry run completed successfully"
    }
    else {
        Write-Step "Publishing to NuGet.org..."
        
        # Confirm before publishing
        Write-Warning "`nAbout to publish:"
        Write-Information "  Package: $packageName" -InformationAction Continue
        Write-Information "  Source:  $NuGetSource" -InformationAction Continue
        
        $confirmation = Read-Host "`nDo you want to continue? (yes/no)"
        
        if ($confirmation -ne "yes") {
            Write-Warning "`n? Publish cancelled by user"
            exit 0
        }

        dotnet nuget push $packagePath.FullName --api-key $apiKey --source $NuGetSource --skip-duplicate
        
        if ($LASTEXITCODE -ne 0) {
            Write-ErrorMessage "Failed to publish package to NuGet"
            exit 1
        }

        Write-Success "Package published successfully to NuGet.org!"
        Write-Information "`n?? Package $packageName is now available on NuGet.org" -InformationAction Continue
        Write-Verbose "   It may take a few minutes to appear in search results.`n" -Verbose
    }

}
catch {
    Write-Error "`n? Error: $_"
    Write-Verbose $_.ScriptStackTrace -Verbose
    exit 1
}
finally {
    # Clean sensitive data from memory
    if ($apiKey) {
        Clear-Variable apiKey -ErrorAction SilentlyContinue
    }
}
