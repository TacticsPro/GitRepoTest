using ClosedXML.Excel;

namespace Office_Tools_Lite.Task_Helper;
public class Excel_Appender
{
    public static readonly string DefaultFirstRowContent = "Goods and Services Tax";

    public static async Task<bool> AppendExcelFiles(IEnumerable<string> inputExcelFiles, string inputSheetName, string outputExcelFile, string outputSheetname, string firstRowContent = null)
    {
        firstRowContent ??= DefaultFirstRowContent; // Use default if null
        using var outputWorkbook = new XLWorkbook();
        var outputWorksheet = outputWorkbook.Worksheets.Add(outputSheetname);
        var currentRow = 1;
        bool allFilesProcessed = true;

        foreach (var excelFile in inputExcelFiles)
        {
            using var inputWorkbook = new XLWorkbook(excelFile);
            var inputWorksheet = inputWorkbook.Worksheets.FirstOrDefault(ws => ws.Name == inputSheetName);
            if (inputWorksheet == null)
            {
                await ShowDialog.ShowMsgBox("Warning", $"'{inputSheetName}' sheet does not exist in selected file {excelFile}", "Ok", null, 1, App.MainWindow);
                allFilesProcessed = false;
                break;
            }

            // Check for the first row contents
            bool isValidSheet = false;
            var firstRow = inputWorksheet.Row(1);
            foreach (var cell in firstRow.CellsUsed())
            {
                string cellValue = cell.GetString().Trim();
                if (!string.IsNullOrEmpty(cellValue) && cellValue.Contains(firstRowContent, StringComparison.OrdinalIgnoreCase))
                {
                    isValidSheet = true;
                    break;
                }
            }

            if (!isValidSheet)
            {
                await ShowDialog.ShowMsgBox("Warning", $"Excel file {excelFile} has invalid data", "Ok", null, 1, App.MainWindow);
                allFilesProcessed = false;
                break;
            }

            var lastCell = inputWorksheet.LastCellUsed();
            if (lastCell == null) continue;

            var startRow = 1;
            var endRow = 40;
            var startCol = 1;
            var endCol = lastCell.Address.ColumnNumber;

            for (var row = startRow; row <= endRow; row++)
            {
                for (var col = startCol; col <= endCol; col++)
                {
                    outputWorksheet.Cell(currentRow, col - startCol + 1).Value = inputWorksheet.Cell(row, col).Value;
                }
                currentRow++;
            }
        }

        if (allFilesProcessed && currentRow > 1) // Only save if data was appended
        {
            outputWorkbook.SaveAs(outputExcelFile);
        }
        return allFilesProcessed;
    }

    #region Add List Sheet
    public static async Task<bool> AddListSheet(string targetExcelFile, string sourceExcelFile, string sourceSheetName, string targetSheetName)
    {
        if (!File.Exists(targetExcelFile))
        {
            await ShowDialog.ShowMsgBox("Error", $"Target file not found: {targetExcelFile}", "Ok", null, 1, App.MainWindow);
            return false;
        }

        if (!File.Exists(sourceExcelFile))
        {
            await ShowDialog.ShowMsgBox("Error", $"Source file not found: {sourceExcelFile}", "Ok", null, 1, App.MainWindow);
            return false;
        }

        try
        {
            using var targetWb = new XLWorkbook(targetExcelFile);
            using var sourceWb = new XLWorkbook(sourceExcelFile);

            var sourceSheet = sourceWb.Worksheets
                .FirstOrDefault(ws => ws.Name.Equals(sourceSheetName, StringComparison.OrdinalIgnoreCase));

            if (sourceSheet == null)
            {
                await ShowDialog.ShowMsgBox("Warning",
                    $"Sheet '{sourceSheetName}' not found in {Path.GetFileName(sourceExcelFile)}", "Ok", null, 1, App.MainWindow);
                return false;
            }

            // Remove existing sheet if present
            if (targetWb.Worksheets.Contains(targetSheetName))
                targetWb.Worksheets.Delete(targetSheetName);

            // Copy sheet
            sourceSheet.CopyTo(targetWb, targetSheetName);
            targetWb.Save();
            return true;
        }
        catch (Exception ex)
        {
            await ShowDialog.ShowMsgBox("Error", $"Failed to add sheet: {ex.Message}", "Ok", null, 1, App.MainWindow);
            return false;
        }
    }
    #endregion
}
