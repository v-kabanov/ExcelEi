@echo off

echo Using %nuget%

rmdir /S /Q "%~dp0NuGet"
%nuget% pack "%~dp0src\ExcelEi" -Symbols -Properties Configuration=Release -OutputDirectory "%~dp0NuGet"