using System;
using System.IO;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace ExcelTestCreator
{
    class Program
    {
        private static Application? excelApp;
        private static Workbook? workbook;

        static void Main(string[] args)
        {
            Console.WriteLine("ThreadSafe Test Workbook Creator");
            Console.WriteLine("=================================");

            try
            {
                // Get the directory containing the CSV files
                string csvDirectory = GetCsvDirectory();
                Console.WriteLine($"CSV Directory: {csvDirectory}");

                // Start Excel
                Console.WriteLine("Starting Excel...");
                StartExcel();

                // Create new workbook
                Console.WriteLine("Creating new workbook...");
                CreateWorkbook();

                // Import each CSV file to its respective sheet
                ImportCsvToSheet(Path.Combine(csvDirectory, "Sheet1_C_Functions_Direct.csv"), "C Functions Direct", 1);
                ImportCsvToSheet(Path.Combine(csvDirectory, "Sheet2_CS_Functions_Direct.csv"), "CS Functions Direct", 2);
                ImportCsvToSheet(Path.Combine(csvDirectory, "Sheet3_Test_Wrappers.csv"), "Test Wrappers", 3);

                // Format the workbook
                FormatWorkbook();

                // Save the workbook
                string outputPath = Path.Combine(csvDirectory, "ThreadSafeTest.xlsx");
                Console.WriteLine($"Saving workbook to: {outputPath}");
                SaveWorkbook(outputPath);

                Console.WriteLine("Workbook created successfully!");
                
                // Only wait for key press if running interactively
                if (Environment.UserInteractive && !Console.IsInputRedirected)
                {
                    Console.WriteLine("Press any key to continue...");
                    Console.ReadKey();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine($"Stack Trace: {ex.StackTrace}");
            }
            finally
            {
                CleanupExcel();
            }
        }

        private static string GetCsvDirectory()
        {
            // Look for CSV files in the current directory or parent directory
            string currentDir = Directory.GetCurrentDirectory();
            
            // Check if we're in the ExcelTestCreator subdirectory
            if (Path.GetFileName(currentDir) == "ExcelTestCreator")
            {
                return Path.GetDirectoryName(currentDir) ?? currentDir;
            }
            
            return currentDir;
        }

        private static void StartExcel()
        {
            excelApp = new Application
            {
                Visible = true,
                DisplayAlerts = false
            };
        }

        private static void CreateWorkbook()
        {
            if (excelApp == null) throw new InvalidOperationException("Excel not started");
            
            workbook = excelApp.Workbooks.Add();
            
            // Ensure we have at least 3 worksheets
            while (workbook.Worksheets.Count < 3)
            {
                workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
            }

            // Rename the worksheets
            ((Worksheet)workbook.Worksheets[1]).Name = "C Functions Direct";
            ((Worksheet)workbook.Worksheets[2]).Name = "CS Functions Direct";
            ((Worksheet)workbook.Worksheets[3]).Name = "Test Wrappers";
        }

        private static void ImportCsvToSheet(string csvPath, string sheetName, int sheetIndex)
        {
            Console.WriteLine($"Importing {csvPath} to sheet '{sheetName}'...");
            
            if (!File.Exists(csvPath))
            {
                Console.WriteLine($"Warning: CSV file not found: {csvPath}");
                return;
            }

            if (workbook == null) throw new InvalidOperationException("Workbook not created");

            Worksheet worksheet = (Worksheet)workbook.Worksheets[sheetIndex];
            
            // Read CSV file
            string[] lines = File.ReadAllLines(csvPath);
            
            for (int row = 0; row < lines.Length; row++)
            {
                string[] columns = ParseCsvLine(lines[row]);
                
                for (int col = 0; col < columns.Length; col++)
                {
                    Range cell = worksheet.Cells[row + 1, col + 1];
                    string cellValue = columns[col].Trim();
                    
                    // Check if this is a formula (starts with =)
                    if (cellValue.StartsWith("="))
                    {
                        // Convert semicolons to commas for US Excel
                        string formula = cellValue.Replace(";", ",");
                        
                        try
                        {
                            // Use Formula2 property with dynamic casting for better compatibility
                            ((dynamic)cell).Formula2 = formula;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Warning: Failed to set formula '{formula}' in cell {GetCellAddress(row + 1, col + 1)}: {ex.Message}");
                            // Fall back to setting as text
                            cell.Value2 = cellValue;
                        }
                    }
                    else if (cellValue.Equals("TRUE", StringComparison.OrdinalIgnoreCase))
                    {
                        cell.Value2 = true;
                    }
                    else if (cellValue.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
                    {
                        cell.Value2 = false;
                    }
                    else if (double.TryParse(cellValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double numValue))
                    {
                        cell.Value2 = numValue;
                    }
                    else if (!string.IsNullOrEmpty(cellValue))
                    {
                        cell.Value2 = cellValue;
                    }
                }
            }

            Console.WriteLine($"Imported {lines.Length} rows to '{sheetName}'");
        }

        private static string[] ParseCsvLine(string line)
        {
            var result = new List<string>();
            bool inQuotes = false;
            string currentField = "";

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    inQuotes = !inQuotes;
                }
                else if (c == ',' && !inQuotes)
                {
                    result.Add(currentField.Trim());
                    currentField = "";
                }
                else
                {
                    currentField += c;
                }
            }

            result.Add(currentField.Trim());
            return result.ToArray();
        }

        private static string GetCellAddress(int row, int col)
        {
            string columnName = "";
            while (col > 0)
            {
                col--;
                columnName = (char)('A' + col % 26) + columnName;
                col /= 26;
            }
            return columnName + row;
        }

        private static void FormatWorkbook()
        {
            if (workbook == null) return;

            Console.WriteLine("Formatting workbook...");

            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                try
                {
                    // Format header row
                    Range headerRow = worksheet.Range["1:1"];
                    headerRow.Font.Bold = true;
                    headerRow.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                    // Auto-fit columns
                    worksheet.Columns.AutoFit();

                    // Set some minimum column widths
                    worksheet.Columns[1].ColumnWidth = 25; // Test Description
                    worksheet.Columns[2].ColumnWidth = 20; // Function
                    worksheet.Columns[5].ColumnWidth = 30; // Results
                    worksheet.Columns[7].ColumnWidth = 25; // Notes

                    // Freeze the header row
                    worksheet.Rows[2].Select();
                    if (excelApp != null && excelApp.ActiveWindow != null)
                    {
                        excelApp.ActiveWindow.FreezePanes = true;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Failed to format worksheet '{worksheet.Name}': {ex.Message}");
                }
            }

            // Select the first worksheet
            ((Worksheet)workbook.Worksheets[1]).Activate();
        }

        private static void SaveWorkbook(string outputPath)
        {
            if (workbook == null) throw new InvalidOperationException("Workbook not created");

            try
            {
                // Delete existing file if it exists
                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }

                workbook.SaveAs(outputPath, XlFileFormat.xlOpenXMLWorkbook);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
                // Try saving with a timestamp
                string timestampPath = Path.Combine(
                    Path.GetDirectoryName(outputPath) ?? "",
                    Path.GetFileNameWithoutExtension(outputPath) + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + Path.GetExtension(outputPath)
                );
                workbook.SaveAs(timestampPath, XlFileFormat.xlOpenXMLWorkbook);
                Console.WriteLine($"Saved as: {timestampPath}");
            }
        }

        private static void CleanupExcel()
        {
            try
            {
                if (workbook != null)
                {
                    workbook.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                }

                if (excelApp != null)
                {
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Error during cleanup: {ex.Message}");
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
    }
}
