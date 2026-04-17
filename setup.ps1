# setup.ps1

Write-Host "Checking for .NET SDK..." -ForegroundColor Cyan

# Check if dotnet is available
if (Get-Command dotnet -ErrorAction SilentlyContinue) {
    $version = dotnet --version
    Write-Host ".NET SDK is already installed (Version: $version)." -ForegroundColor Green
} else {
    Write-Host ".NET SDK not found. Installing .NET 8 SDK..." -ForegroundColor Yellow
    
    # Use winget to install the .NET 8 SDK silently
    winget install Microsoft.DotNet.SDK.8 --accept-package-agreements --accept-source-agreements
    
    Write-Host ""
    Write-Host "========================================================" -ForegroundColor Yellow
    Write-Host "Installation complete!" -ForegroundColor Green
    Write-Host "IMPORTANT: You must restart your terminal (or VS Code/Cursor)" -ForegroundColor Red
    Write-Host "so that the 'dotnet' command is recognized in your PATH." -ForegroundColor Red
    Write-Host "After restarting, run this script again." -ForegroundColor Yellow
    Write-Host "========================================================" -ForegroundColor Yellow
    exit
}

Write-Host "`nBuilding the project..." -ForegroundColor Cyan
dotnet build src/VbaMacroParser/VbaMacroParser.csproj

Write-Host "`nRunning tests..." -ForegroundColor Cyan
dotnet test tests/VbaMacroParser.Tests/VbaMacroParser.Tests.csproj

Write-Host "`nSetup and verification complete! You can now run the app using:" -ForegroundColor Green
Write-Host "dotnet run --project src/VbaMacroParser/VbaMacroParser.csproj" -ForegroundColor White