@echo off
echo Checking for .NET SDK...

:: Check if dotnet is available
where dotnet >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo .NET SDK not found. Installing .NET 8 SDK...
    
    :: Use winget to install the .NET 8 SDK
    winget install Microsoft.DotNet.SDK.8 --accept-package-agreements --accept-source-agreements
    
    echo.
    echo ========================================================
    echo Installation complete!
    echo IMPORTANT: You must restart your terminal ^(or VS Code/Cursor^)
    echo so that the 'dotnet' command is recognized in your PATH.
    echo After restarting, run this script again.
    echo ========================================================
    pause
    exit /b
) else (
    echo .NET SDK is already installed.
)

echo.
echo Building the project...
dotnet build src/VbaMacroParser/VbaMacroParser.csproj

echo.
echo Running tests...
dotnet test tests/VbaMacroParser.Tests/VbaMacroParser.Tests.csproj

echo.
echo Setup and verification complete! You can now run the app using either of the following commands:
echo 1. ./run.bat
echo 2. dotnet run --project src/VbaMacroParser/VbaMacroParser.csproj