@echo off
echo ThreadSafe Test Workbook Creator
echo ===================================
echo.
echo This script will create a comprehensive Excel test workbook
echo from the CSV templates using the ExcelTestCreator C# console application.
echo.
echo Prerequisites:
echo - .NET 9 runtime installed
echo - Excel installed on the system
echo - All CSV template files present in this directory
echo.
pause

cd ExcelTestCreator
dotnet run
cd ..

echo.
echo Done! Check ThreadSafeTest.xlsx in this directory.
pause