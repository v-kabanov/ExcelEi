@echo off

if not defined PACKOUT (
    set PACKOUT=%~dp0NuGet
)

if not exist "%PACKOUT%" (
    echo Creating "%PACKOUT%"
    mkdir "%PACKOUT%"
)

echo Packing into %PACKOUT%

dotnet pack "%~dp0src\ExcelEi\ExcelEi.csproj" -c Release --include-symbols --include-source -o "%PACKOUT%"
