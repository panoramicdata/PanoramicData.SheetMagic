# Publishing to NuGet

This document explains how to publish the PanoramicData.SheetMagic package to NuGet.org.

## Prerequisites

1. **PowerShell 7.0+** - The publish script requires PowerShell 7 or later
2. **NuGet API Key** - You need a valid NuGet.org API key
3. **Git** - A clean working directory (no uncommitted changes)
4. **.NET 9 SDK** - To build and test the project

## Setup

### 1. Get Your NuGet API Key

1. Go to [https://www.nuget.org/account/apikeys](https://www.nuget.org/account/apikeys)
2. Sign in with your NuGet.org account
3. Create a new API key with the following settings:
   - **Key Name**: `PanoramicData.SheetMagic`
   - **Package Owner**: Select your account
   - **Scopes**: Select "Push" and "Push new packages and package versions"
   - **Packages**: Select "Only selected packages" and choose `PanoramicData.SheetMagic`
   - **Expiration**: Set an appropriate expiration date

### 2. Store Your API Key

Add your API key to the `nuget-key.txt` file in the root directory:

```
your-api-key-here
```

**?? IMPORTANT:** This file is already in `.gitignore`. Never commit your API key to source control!

## Publishing

### Standard Publish

To publish with all safety checks:

```powershell
.\Publish.ps1
```

This will:
1. ? Check Git status (must be clean)
2. ? Run all unit tests
3. ? Build and pack the project
4. ? Prompt for confirmation
5. ? Publish to NuGet.org

### Dry Run

To test the script without actually publishing:

```powershell
.\Publish.ps1 -DryRun
```

### Skip Tests (Not Recommended)

To skip unit tests (use only in exceptional circumstances):

```powershell
.\Publish.ps1 -SkipTests
```

## What the Script Does

### 1. Git Status Check
Ensures your working directory is clean with no uncommitted changes. This guarantees that the published package matches a specific commit in your repository.

### 2. Unit Tests
Runs all tests in Release configuration to ensure the package works correctly.

### 3. Build & Pack
Builds the project in Release configuration and creates the NuGet package (.nupkg file).

### 4. Publish
Pushes the package to NuGet.org. The `--skip-duplicate` flag prevents errors if the version already exists.

## Versioning

The project uses [Nerdbank.GitVersioning](https://github.com/dotnet/Nerdbank.GitVersioning) for automatic version management. The version is determined by:

- Git tags
- Commit height
- Branch name
- Configuration in `version.json`

### Creating a Release

1. **Create a version tag:**
   ```bash
   git tag v1.2.3
   git push origin v1.2.3
   ```

2. **Ensure clean state:**
   ```bash
   git status
   ```

3. **Publish:**
   ```powershell
   .\Publish.ps1
   ```

## Troubleshooting

### "Git working directory is not clean"
- Commit or stash your changes before publishing
- Use `git status` to see what's uncommitted

### "Unit tests failed"
- Fix the failing tests before publishing
- Review test output for details
- Use `-SkipTests` only if absolutely necessary (not recommended)

### "NuGet API key file not found"
- Ensure `nuget-key.txt` exists in the root directory
- Add your API key to the file

### "Failed to publish package"
- Check if the version already exists on NuGet.org
- Verify your API key has the correct permissions
- Ensure you have internet connectivity

### "Could not find generated NuGet package"
- Check the build output for errors
- Ensure the project builds successfully
- Look in `PanoramicData.SheetMagic\bin\Release` for .nupkg files

## Security Notes

- ? `nuget-key.txt` is in `.gitignore`
- ? The script clears the API key from memory after use
- ? Never commit API keys to source control
- ? Rotate API keys periodically
- ? Use scoped API keys (only for this package)

## Best Practices

1. **Always** ensure tests pass before publishing
2. **Always** commit your changes before publishing
3. **Tag** releases for version tracking
4. **Update** package release notes in the .csproj file
5. **Test** in a staging environment if possible
6. **Review** the package contents before publishing
7. **Monitor** NuGet.org after publishing to ensure the package appears

## Additional Resources

- [NuGet Package Publishing Guide](https://docs.microsoft.com/en-us/nuget/nuget-org/publish-a-package)
- [Nerdbank.GitVersioning Documentation](https://github.com/dotnet/Nerdbank.GitVersioning/blob/main/doc/nbgv-cli.md)
- [Semantic Versioning](https://semver.org/)
