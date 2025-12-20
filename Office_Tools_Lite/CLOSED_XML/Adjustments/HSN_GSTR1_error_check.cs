using System.Diagnostics;
using System.IO.Compression;
using System.Text.Json;
using ClosedXML.Excel;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Office_Tools_Lite.Task_Helper;
using Windows.Storage;

namespace Office_Tools_Lite.CLOSED_XML.Adjustments;
public class HSN_GSTR1_error_check
{
    #region Execute Main Task
    public async Task Execute(Window mainWindow)
    {
        var dialog = await ShowDialog.ShowMsgBox(
            "GSTR-1 Error Check",
            "Select downloaded error .zip file from GST website.\n\n Click OK to continue.",
            "OK", "Cancel", 1, mainWindow);

        if (dialog != ContentDialogResult.Primary)
        {
            return;
        }

        var file = await Filepick();
        if (file == null)
        {
            return;
        }

        var zipFilePath = file.Path;
        var outputDirectory = Path.GetDirectoryName(zipFilePath);
        var excelFilePath = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(zipFilePath) + ".xlsx");

        await ReadjsonFileAndWriteExcel(zipFilePath, excelFilePath);

    }
    #endregion

    #region Filepicker
    private async Task<StorageFile?> Filepick()
    {
        // Open File Picker to select a single Excel file
        var picker = new Windows.Storage.Pickers.FileOpenPicker();
        picker.FileTypeFilter.Add(".zip");
        picker.SuggestedStartLocation = Windows.Storage.Pickers.PickerLocationId.DocumentsLibrary;

        var hwnd = WinRT.Interop.WindowNative.GetWindowHandle(App.MainWindow);
        WinRT.Interop.InitializeWithWindow.Initialize(picker, hwnd);

        var file = await picker.PickSingleFileAsync();

        if (file == null)
        {
            await ShowDialog.ShowMsgBox("Warning", "No Zip file were selected.", "OK", null, 1, App.MainWindow);
            return null;
        }
        return file;
    }
    #endregion

    #region Read Json File and Write Excel
    private async Task ReadjsonFileAndWriteExcel(string zipFilePath, string excelFilePath)
    {
        try
        {
            // Create Excel workbook
            using (var workbook = new XLWorkbook())
            {
                bool hasData = false;

                // Extract JSON files from ZIP
                using (ZipArchive archive = ZipFile.OpenRead(zipFilePath))
                {
                    foreach (ZipArchiveEntry entry in archive.Entries)
                    {
                        if (!entry.FullName.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
                            continue;

                        // Read JSON content
                        string jsonContent;
                        using (StreamReader reader = new StreamReader(entry.Open()))
                        {
                            jsonContent = reader.ReadToEnd();
                        }

                        if (string.IsNullOrEmpty(jsonContent))
                            continue;

                        // Parse JSON
                        using (JsonDocument doc = JsonDocument.Parse(jsonContent))
                        {
                            JsonElement root = doc.RootElement;
                            if (!root.TryGetProperty("error_report", out JsonElement errorReport))
                                continue;

                            // Process all error sections dynamically
                            foreach (var section in errorReport.EnumerateObject())
                            {
                                if (section.Value.ValueKind == JsonValueKind.Array)
                                {
                                    ProcessErrorSection(section.Name, section.Value, workbook);
                                    hasData = true;
                                }
                            }
                        }
                    }
                }

                if (!hasData)
                {
                    await ShowDialog.ShowMsgBox("Information", "No valid error data found in the JSON files.", "OK", null, 1, App.MainWindow);
                    return;
                }

                // Save the Excel file
                workbook.SaveAs(excelFilePath);
                await ShowDialog.ShowMsgBox("Success", $"Excel file saved successfully at: {excelFilePath}", "OK", null, 1, App.MainWindow);
                try
                {
                    Process.Start(new ProcessStartInfo(excelFilePath) { UseShellExecute = true });
                }
                catch (Exception ex)
                {
                    await ShowDialog.ShowMsgBox("Error", $"Failed to open the Excel file: {ex.Message}", "OK", null, 1, App.MainWindow);
                }
            }
        }
        catch (Exception ex)
        {
            await ShowDialog.ShowMsgBox("Error", $"An error occurred: {ex.Message}", "OK", null, 1, App.MainWindow);
        }

    }
    #endregion

    #region Process json file and Highlight error
    private void ProcessErrorSection(string sectionName, JsonElement errorArray, XLWorkbook workbook)
    {
        // Sanitize section name for worksheet
        string worksheetName = sectionName.Length > 31 ? sectionName.Substring(0, 31) : sectionName;
        worksheetName = worksheetName.Replace(":", "").Replace("/", "").Replace("\\", "").Replace("?", "").Replace("*", "").Replace("[", "").Replace("]", "");
        var worksheet = workbook.Worksheets.Add($"{worksheetName} Errors");

        // Collect headers and rows
        var headers = new HashSet<string>();
        var rows = new List<Dictionary<string, JsonElement>>();

        if (errorArray.EnumerateArray().Any())
        {
            foreach (JsonElement item in errorArray.EnumerateArray())
            {
                var currentRow = new Dictionary<string, JsonElement>();
                FlattenJson(item, "", currentRow, rows, headers);
            }
        }

        var headerList = headers.ToList();

        // Set headers in Excel
        for (int i = 0; i < headerList.Count; i++)
        {
            worksheet.Cell(1, i + 1).Value = headerList[i];
        }

        // Write data
        int row = 2;
        var seenHsnKeys = new HashSet<string>();

        foreach (var rowData in rows)
        {
            for (int i = 0; i < headerList.Count; i++)
            {
                string header = headerList[i];
                if (rowData.TryGetValue(header, out JsonElement value))
                {
                    WriteJsonValue(worksheet, row, i + 1, value);
                }
            }

            if (sectionName.Equals("hsn", StringComparison.OrdinalIgnoreCase))
            {
                // --- Highlight error_msg rows (Yellow) ---
                if (headerList.Contains("error_msg"))
                {
                    int errCol = headerList.IndexOf("error_msg") + 1;

                    for (int r = 2; r < row; r++)
                    {
                        string errMsg = worksheet.Cell(r, errCol).GetString();
                        if (!string.IsNullOrWhiteSpace(errMsg))
                        {
                            worksheet.Row(r).Style.Fill.BackgroundColor = XLColor.Yellow;
                        }
                    }
                }

                // --- Duplicate check (HSN, UQC, Rate) ---
                if (headerList.Contains("hsn_b2c.hsn_sc") &&
                    headerList.Contains("hsn_b2c.uqc") &&
                    headerList.Contains("hsn_b2c.rt"))
                {
                    var hsnGroups = new Dictionary<string, List<int>>();

                    int hsnCol = headerList.IndexOf("hsn_b2c.hsn_sc") + 1;
                    int uqcCol = headerList.IndexOf("hsn_b2c.uqc") + 1;
                    int rtCol = headerList.IndexOf("hsn_b2c.rt") + 1;

                    for (int r = 2; r < row; r++)
                    {
                        string hsn = worksheet.Cell(r, hsnCol).GetString();
                        string uqc = worksheet.Cell(r, uqcCol).GetString();
                        string rt = worksheet.Cell(r, rtCol).GetString();

                        string key = $"{hsn}|{uqc}|{rt}";

                        if (!hsnGroups.ContainsKey(key))
                            hsnGroups[key] = new List<int>();

                        hsnGroups[key].Add(r);
                    }

                    foreach (var group in hsnGroups.Values)
                    {
                        if (group.Count > 1)
                        {
                            foreach (var r in group)
                            {
                                worksheet.Row(r).Style.Fill.BackgroundColor = XLColor.LightPink;
                            }
                        }
                    }
                }
            }


            row++;
        }

        worksheet.Columns().AdjustToContents();
    }
    #endregion

    #region Get json Headers
    private void FlattenJson(JsonElement element, string prefix, Dictionary<string, JsonElement> currentRow, List<Dictionary<string, JsonElement>> rows, HashSet<string> headers)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            // First pass: add all primitive fields and detect if has sub child
            bool hasSub = false;
            foreach (var prop in element.EnumerateObject())
            {
                string newPrefix = string.IsNullOrEmpty(prefix) ? prop.Name : $"{prefix}.{prop.Name}";
                if (prop.Value.ValueKind == JsonValueKind.Array || prop.Value.ValueKind == JsonValueKind.Object)
                {
                    hasSub = true;
                }
                else
                {
                    if (!headers.Contains(newPrefix))
                        headers.Add(newPrefix);
                    currentRow[newPrefix] = prop.Value;
                }
            }

            // Second pass: process non-primitives
            foreach (var prop in element.EnumerateObject())
            {
                string newPrefix = string.IsNullOrEmpty(prefix) ? prop.Name : $"{prefix}.{prop.Name}";
                if (prop.Value.ValueKind == JsonValueKind.Array)
                {
                    foreach (var arrayItem in prop.Value.EnumerateArray())
                    {
                        var clonedRow = new Dictionary<string, JsonElement>(currentRow);
                        FlattenJson(arrayItem, newPrefix, clonedRow, rows, headers);
                    }
                }
                else if (prop.Value.ValueKind == JsonValueKind.Object)
                {
                    FlattenJson(prop.Value, newPrefix, currentRow, rows, headers);
                }
            }

            if (!hasSub)
            {
                rows.Add(new Dictionary<string, JsonElement>(currentRow));
            }
        }
        else if (element.ValueKind == JsonValueKind.Array)
        {
            foreach (var arrayItem in element.EnumerateArray())
            {
                var clonedRow = new Dictionary<string, JsonElement>(currentRow);
                FlattenJson(arrayItem, prefix, clonedRow, rows, headers);
            }
        }
        else
        {
            // Primitive at root, rare
            string newPrefix = prefix;
            if (!headers.Contains(newPrefix))
                headers.Add(newPrefix);
            currentRow[newPrefix] = element;
        }
    }
    #endregion

    #region Write json Values
    private void WriteJsonValue(IXLWorksheet worksheet, int row, int col, JsonElement value)
    {
        switch (value.ValueKind)
        {
            case JsonValueKind.String:
                worksheet.Cell(row, col).Value = value.GetString();
                break;
            case JsonValueKind.Number:
                if (value.TryGetDouble(out double doubleValue))
                    worksheet.Cell(row, col).Value = doubleValue;
                else if (value.TryGetInt32(out int intValue))
                    worksheet.Cell(row, col).Value = intValue;
                break;
            case JsonValueKind.True:
                worksheet.Cell(row, col).Value = true;
                break;
            case JsonValueKind.False:
                worksheet.Cell(row, col).Value = false;
                break;
            case JsonValueKind.Null:
                worksheet.Cell(row, col).Value = string.Empty;
                break;
        }
    }
    #endregion
}
