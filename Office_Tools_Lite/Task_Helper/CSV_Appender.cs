using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace Office_Tools_Lite.Task_Helper;
public class CSV_Appender
{
    public static readonly string DefaultFirstRowContent = "Goods and Services Tax";

    public static async Task<bool> AppendCsvToExcel(IEnumerable<string> inputcsvFiles, string outputExcelFile, string outputSheetName, string firstRowContent = null)
    {
        firstRowContent ??= DefaultFirstRowContent; // Use default if null
        using var outputWorkbook = new XLWorkbook();
        var outputWorksheet = outputWorkbook.Worksheets.Add(outputSheetName);
        var currentRow = 1;
        bool allFilesProcessed = true;

        foreach (var csvFile in inputcsvFiles)
        {
            bool isValidFile = false;
            bool isFirstRow = true;

            // Process the CSV files using their paths
            using var reader = new StreamReader(csvFile);
            while (!reader.EndOfStream)
            {
                var line = await reader.ReadLineAsync();
                if (string.IsNullOrWhiteSpace(line)) continue; // Skip empty lines

                var columns = line.Split(',');

                // Check for the first row contents
                if (isFirstRow)
                {
                    foreach (var column in columns)
                    {
                        var cellValue = Regex.Replace(column, @"\s+", " ").Trim('\"').Trim();
                        if (!string.IsNullOrEmpty(cellValue) && cellValue.Contains(firstRowContent, StringComparison.OrdinalIgnoreCase))
                        {
                            isValidFile = true;
                            break;
                        }
                    }

                    if (!isValidFile)
                    {
                        await ShowDialog.ShowMsgBox("Warning", $"CSV file {csvFile} has invalid data", "Ok", null, 1, App.MainWindow);
                        allFilesProcessed = false;
                        break; // Stop processing this and further files
                    }
                }

                for (var col = 0; col < columns.Length; col++)
                {
                    //var cellValue = Regex.Replace(columns[col], @"\s+", " ").Trim('"').Trim(); //Removes spaces
                    var cellValue = Regex.Replace(columns[col], @"\s+", " ").Trim('\"');

                    // Force string for non-numeric ID columns (e.g., invoice number at col=1)
                    if (col == 1 || col == 2)  // Invoice number and CN-DN note number column
                    {
                        if (cellValue.Length > 15)
                        {
                            // Long IDs: Force as text to preserve exact digits
                            outputWorksheet.Cell(currentRow, col + 1).Value = cellValue;
                            outputWorksheet.Cell(currentRow, col + 1).Style.NumberFormat.Format = "@";  // Text format
                        }
                        else
                        {
                            // Short values: Try numeric, fallback to string
                            if (double.TryParse(cellValue, out var numericValue))
                            {
                                outputWorksheet.Cell(currentRow, col + 1).Value = numericValue;
                            }
                            else
                            {
                                outputWorksheet.Cell(currentRow, col + 1).Value = cellValue;
                            }
                        }
                    }
                    else if (double.TryParse(cellValue, out var numericValue))
                    {
                        outputWorksheet.Cell(currentRow, col + 1).Value = numericValue;
                    }
                    else
                    {
                        outputWorksheet.Cell(currentRow, col + 1).Value = cellValue;
                    }
                }

                currentRow++;
                if (currentRow == 40)
                {
                    break;
                }
            }
        }

        if (allFilesProcessed && currentRow > 1) // Only save if data was appended
        {
            outputWorkbook.SaveAs(outputExcelFile);
        }
        return allFilesProcessed;
    }
}
