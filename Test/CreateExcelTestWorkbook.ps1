# ThreadSafe Test Workbook Creator PowerShell Script
# Creates Excel workbook from CSV templates using COM automation

param(
    [string]$OutputPath = "ThreadSafeTest.xlsx"
)

Write-Host "ThreadSafe Test Workbook Creator (PowerShell)" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green

try {
    # Get current script directory
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    Write-Host "Working directory: $scriptDir" -ForegroundColor Yellow

    # Define CSV files
    $csvFiles = @{
        "Sheet1_C_Functions_Direct.csv" = "C Functions Direct"
        "Sheet2_CS_Functions_Direct.csv" = "CS Functions Direct"  
        "Sheet3_Test_Wrappers.csv" = "Test Wrappers"
    }

    # Create Excel application
    Write-Host "Starting Excel..." -ForegroundColor Yellow
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false

    # Create new workbook
    Write-Host "Creating new workbook..." -ForegroundColor Yellow
    $workbook = $excel.Workbooks.Add()

    # Ensure we have 3 worksheets
    while ($workbook.Worksheets.Count -lt 3) {
        $workbook.Worksheets.Add() | Out-Null
    }

    $sheetIndex = 1
    foreach ($csvFile in $csvFiles.Keys) {
        $csvPath = Join-Path $scriptDir $csvFile
        $sheetName = $csvFiles[$csvFile]
        
        Write-Host "Processing $csvFile -> '$sheetName'..." -ForegroundColor Yellow
        
        if (-not (Test-Path $csvPath)) {
            Write-Host "Warning: CSV file not found: $csvPath" -ForegroundColor Red
            $sheetIndex++
            continue
        }

        # Get the worksheet and rename it
        $worksheet = $workbook.Worksheets.Item($sheetIndex)
        $worksheet.Name = $sheetName

        # Read CSV file
        $csvData = Import-Csv $csvPath -Header @("Col1","Col2","Col3","Col4","Col5","Col6","Col7")
        
        # Process header row first
        $headerLine = Get-Content $csvPath | Select-Object -First 1
        $headers = $headerLine -split ','
        
        for ($col = 1; $col -le $headers.Length; $col++) {
            $headerValue = $headers[$col-1].Trim('"')
            $worksheet.Cells.Item(1, $col).Value2 = $headerValue
        }
        
        # Process data rows
        $csvLines = Get-Content $csvPath | Select-Object -Skip 1
        $row = 2
        
        foreach ($line in $csvLines) {
            # Parse CSV line properly handling quotes
            $fields = @()
            $currentField = ""
            $inQuotes = $false
            
            for ($i = 0; $i -lt $line.Length; $i++) {
                $char = $line[$i]
                if ($char -eq '"') {
                    $inQuotes = -not $inQuotes
                } elseif ($char -eq ',' -and -not $inQuotes) {
                    $fields += $currentField.Trim()
                    $currentField = ""
                } else {
                    $currentField += $char
                }
            }
            $fields += $currentField.Trim()
            
            # Set cell values
            for ($col = 1; $col -le $fields.Length; $col++) {
                $cellValue = $fields[$col-1]
                $cell = $worksheet.Cells.Item($row, $col)
                
                if ($cellValue.StartsWith("=")) {
                    # This is a formula - convert semicolons to commas for US Excel
                    $formula = $cellValue -replace ";", ","
                    try {
                        $cell.Formula2 = $formula
                        Write-Host "Set formula: $formula" -ForegroundColor Cyan
                    } catch {
                        Write-Host "Warning: Failed to set formula '$formula' in cell $($worksheet.Cells.Item($row, $col).Address()): $($_.Exception.Message)" -ForegroundColor Red
                        $cell.Value2 = $cellValue
                    }
                } elseif ($cellValue -eq "TRUE") {
                    $cell.Value2 = $true
                } elseif ($cellValue -eq "FALSE") {
                    $cell.Value2 = $false
                } elseif ([double]::TryParse($cellValue, [ref]$null)) {
                    $cell.Value2 = [double]$cellValue
                } elseif ($cellValue -ne "") {
                    $cell.Value2 = $cellValue
                }
            }
            $row++
        }
        
        Write-Host "Imported $($csvLines.Count + 1) rows to '$sheetName'" -ForegroundColor Green
        $sheetIndex++
    }

    # Format the workbook
    Write-Host "Formatting workbook..." -ForegroundColor Yellow
    
    foreach ($ws in $workbook.Worksheets) {
        try {
            # Format header row
            $headerRange = $ws.Range("1:1")
            $headerRange.Font.Bold = $true
            $headerRange.Interior.Color = 12632256  # Light gray

            # Auto-fit columns
            $ws.Columns.AutoFit() | Out-Null

            # Set minimum column widths
            $ws.Columns.Item(1).ColumnWidth = 25  # Test Description
            $ws.Columns.Item(2).ColumnWidth = 20  # Function  
            $ws.Columns.Item(5).ColumnWidth = 30  # Results
            if ($ws.Columns.Count -ge 7) {
                $ws.Columns.Item(7).ColumnWidth = 25  # Notes
            }

            # Freeze header row
            $ws.Rows.Item(2).Select() | Out-Null
            $excel.ActiveWindow.FreezePanes = $true
            
            Write-Host "Formatted worksheet: $($ws.Name)" -ForegroundColor Cyan
        } catch {
            Write-Host "Warning: Failed to format worksheet '$($ws.Name)': $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Select first worksheet
    $workbook.Worksheets.Item(1).Activate() | Out-Null

    # Save the workbook
    $outputFullPath = Join-Path $scriptDir $OutputPath
    Write-Host "Saving workbook to: $outputFullPath" -ForegroundColor Yellow

    try {
        if (Test-Path $outputFullPath) {
            Remove-Item $outputFullPath -Force
        }
        $workbook.SaveAs($outputFullPath, 51)  # xlOpenXMLWorkbook = 51
        Write-Host "Workbook created successfully!" -ForegroundColor Green
    } catch {
        # Try with timestamp if save fails
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $timestampPath = Join-Path $scriptDir "ThreadSafeTest_$timestamp.xlsx"
        $workbook.SaveAs($timestampPath, 51)
        Write-Host "Saved as: $timestampPath" -ForegroundColor Green
    }

} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.Exception.StackTrace -ForegroundColor Red
} finally {
    # Clean up Excel COM objects
    if ($workbook) {
        $workbook.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
    
    Write-Host "Press any key to continue..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}