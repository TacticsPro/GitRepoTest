using System.Text;
using ClosedXML.Excel;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Office_Tools_Lite.Task_Helper;
using Windows.Storage.Pickers;
using WinRT.Interop;
using static Office_Tools_Lite.Task_Helper.HSN_CSV_Model_Helpers;

namespace Office_Tools_Lite.CLOSED_XML.Adjustments;

public class HSN_Descriptions_Fill
{
    #region Execute Main operation

    public async Task Execute(Window mainWindow)
    {
        var dialog = await ShowDialog.ShowMsgBox(
            "HSN Description Fill",
            "Click OK to continue.",
            "OK", "Cancel", 1, mainWindow);

        if (dialog != ContentDialogResult.Primary)
        {
            return;
        }

        // File picker dialog
        var filePicker = new FileOpenPicker();
        var hwnd = WindowNative.GetWindowHandle(mainWindow);
        InitializeWithWindow.Initialize(filePicker, hwnd);

        filePicker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
        filePicker.FileTypeFilter.Add(".xlsx");
        filePicker.FileTypeFilter.Add(".csv");

        var files = await filePicker.PickSingleFileAsync();
        if (files == null)
        {
            await ShowDialog.ShowMsgBox("Warning", "files were selected", "OK", null, 1, App.MainWindow);
            return;
        }

        var inputExcelFiles = files;
        var outputExcelFile = inputExcelFiles;


        await RunMacro(mainWindow, outputExcelFile);

    }
    #endregion

    #region Run macro
    private async Task RunMacro(Window mainWindow, Windows.Storage.StorageFile inputFile)
    {
        try
        {
            var (validHSNCodes, hsnDescMap, _) = await HSN_CSV_Model_Helpers.LoadValidCodesAsync();
            if (hsnDescMap == null)
            {
                await ShowDialog.ShowMsgBox("Error", "Could not load HSN master data (HSN.json).", "OK", null, 1, mainWindow);
                return;
            }

            string filePath = inputFile.Path;
            string ext = System.IO.Path.GetExtension(filePath).ToLowerInvariant();

            if (ext == ".xlsx")
            {
                using var workbook = new XLWorkbook(filePath);
                var ws = workbook.Worksheets.First();

                var headerRow = ws.FirstRowUsed();
                int hsnCol = -1, descCol = -1;

                // Find HSN and Description column indices
                foreach (var cell in headerRow.CellsUsed())
                {
                    string header = cell.Value.ToString().Trim();
                    if (string.Equals(header, "HSN", StringComparison.OrdinalIgnoreCase))
                        hsnCol = cell.Address.ColumnNumber;
                    else if (string.Equals(header, "Description", StringComparison.OrdinalIgnoreCase))
                        descCol = cell.Address.ColumnNumber;
                }

                if (hsnCol == -1)
                {
                    await ShowDialog.ShowMsgBox("Error", "HSN column not found.", "OK", null, 1, mainWindow);
                    return;
                }
                if (descCol == -1)
                {
                    await ShowDialog.ShowMsgBox("Error", "Description column not found.", "OK", null, 1, mainWindow);
                    return;
                }

                // Iterate through data rows (skip header)
                var dataRows = ws.RowsUsed().Skip(1);
                int updatedCount = 0;
                foreach (var row in dataRows)
                {
                    var hsnCell = row.Cell(hsnCol);
                    var descCell = row.Cell(descCol);

                    // Skip if HSN cell is empty
                    if (hsnCell.IsEmpty() || string.IsNullOrWhiteSpace(hsnCell.GetString()))
                        continue;

                    string hsn = hsnCell.GetString().Trim();

                    // Only update if Description is empty or whitespace
                    if (string.IsNullOrWhiteSpace(descCell.GetString()) &&
                        hsnDescMap.TryGetValue(hsn, out string description) &&
                        !string.IsNullOrWhiteSpace(description))
                    {
                        descCell.Value = description;
                        updatedCount++;
                    }
                }

                workbook.Save(); // ← SAVES TO SAME FILE
            }

            else if (ext == ".csv")
            {
                var lines = await File.ReadAllLinesAsync(filePath);
                if (lines.Length < 2)
                {
                    await ShowDialog.ShowMsgBox("Warning", "CSV file has no data rows.", "OK", null, 1, mainWindow);
                    return;
                }

                var headerLine = lines[0];
                var headerFields = CsvManualParser.ParseCsvLine(headerLine);
                int hsnCol = -1, descCol = -1;

                // Find HSN and Description column indices
                for (int i = 0; i < headerFields.Length; i++)
                {
                    string header = headerFields[i].Trim();
                    if (string.Equals(header, "HSN", StringComparison.OrdinalIgnoreCase))
                        hsnCol = i;
                    else if (string.Equals(header, "Description", StringComparison.OrdinalIgnoreCase))
                        descCol = i;
                }

                if (hsnCol == -1)
                {
                    await ShowDialog.ShowMsgBox("Error", "HSN column not found in CSV.", "OK", null, 1, mainWindow);
                    return;
                }
                if (descCol == -1)
                {
                    await ShowDialog.ShowMsgBox("Error", "Description column not found in CSV.", "OK", null, 1, mainWindow);
                    return;
                }

                var updatedLines = new List<string> { headerLine };
                int updatedCount = 0;

                for (int i = 1; i < lines.Length; i++)
                {
                    var currentLine = lines[i];
                    var fields = CsvManualParser.ParseCsvLine(currentLine);

                    if (fields.Length == 0 || fields.Length <= Math.Max(hsnCol, descCol))
                        continue; // Skip malformed lines

                    var hsnField = fields[hsnCol].Trim();
                    var descField = fields[descCol].Trim();

                    // Skip if HSN is empty
                    if (string.IsNullOrWhiteSpace(hsnField))
                    {
                        updatedLines.Add(currentLine);
                        continue;
                    }

                    // Only update if Description is empty or whitespace
                    if (string.IsNullOrWhiteSpace(descField) &&
                        hsnDescMap.TryGetValue(hsnField, out string description) &&
                        !string.IsNullOrWhiteSpace(description))
                    {
                        // Replace the description field
                        fields[descCol] = description;

                        // Rebuild the line with updated fields
                        var updatedLine = string.Join(",", fields.Select(f => QuoteIfNeeded(f)));
                        updatedLines.Add(updatedLine);
                        updatedCount++;
                    }
                    else
                    {
                        // No change, add original line
                        updatedLines.Add(currentLine);
                    }
                }

                // Write back to file
                await File.WriteAllLinesAsync(filePath, updatedLines, Encoding.UTF8);
                
            }

            await ShowDialog.ShowMsgBox("Success", $"Descriptions updated in: {inputFile.Path}", "OK", null, 1, mainWindow);

            System.Diagnostics.Process.Start("explorer.exe",System.IO.Path.GetDirectoryName(filePath));
        }
        catch (Exception ex)
        {
            await ShowDialog.ShowMsgBox("Error", $"Failed to update file: {ex.Message}", "OK", null, 1, mainWindow);
        }
    }

    private static string QuoteIfNeeded(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        bool needsQuotes = value.Contains(',') || value.Contains('"') || value.Contains('\n') || value.Contains('\r');
        if (!needsQuotes)
            return value;

        // Escape quotes and wrap in double quotes
        string escaped = value.Replace("\"", "\"\"");
        return $"\"{escaped}\"";
    }
    #endregion

}
