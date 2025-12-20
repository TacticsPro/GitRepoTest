using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Office_Tools_Lite.Task_Helper;
using Windows.Storage.Pickers;
using WinRT.Interop;
using ClosedXML.Excel;

namespace Office_Tools_Lite.CLOSED_XML.Adjustments;
public class HSN_json_with_adjust
{
    #region Execute main Task
    public async Task Execute(Window mainWindow)
    {
        await ShowDialog.ShowMsgBox("Notice", "Lite version can only process max 10 row data.\nYou can upgrade from Home Page for Full Feature!\n\nPress OK to Continue with Lite.", "OK", null, 1, App.MainWindow);

        var b2b_b2c = await ShowDialog.ShowMsgBox(
            "Select B2B/B2C (HSN-Direct)",
            "", "B2B", "B2C", 0, mainWindow);

        var result = b2b_b2c;
        string HSNType = result == ContentDialogResult.Primary ? "hsn_b2b" : result == ContentDialogResult.Secondary ? "hsn_b2c" : "";

        if (string.IsNullOrEmpty(HSNType))
        {
            return;
        }

        var filePicker = new FileOpenPicker();
        var hwnd = WindowNative.GetWindowHandle(mainWindow);
        InitializeWithWindow.Initialize(filePicker, hwnd);

        filePicker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
        filePicker.FileTypeFilter.Add(".xlsx");
        filePicker.FileTypeFilter.Add(".csv");

        var files = await filePicker.PickMultipleFilesAsync();
        if (files.Count == 0)
        {
            await ShowDialog.ShowMsgBox("Warning", "No files were selected.", "OK", null, 1, App.MainWindow);
            return;
        }

        var excel_csv_files = files.Select(f => f.Path).ToArray();
        var originalDir = Path.GetDirectoryName(excel_csv_files[0]);
        var subDir = Path.Combine(originalDir, "Output_Files_csv");
        Directory.CreateDirectory(subDir);
        var outputCsvFile = "";

        if (HSNType == "hsn_b2b")
        {
            outputCsvFile = Path.Combine(subDir, "HSN_B2B.csv");
        }
        else if (HSNType == "hsn_b2c")
        {
            outputCsvFile = Path.Combine(subDir, "HSN_B2C.csv");
        }
        var inputExcelCsvFile = excel_csv_files[0];

        await Renamesheet (mainWindow, inputExcelCsvFile);
        await RunMacro(mainWindow, inputExcelCsvFile, outputCsvFile);
        await HSN_json_generator.ConvertCsvToJson(mainWindow, outputCsvFile, subDir, HSNType);
    }
    #endregion

    #region Excel sheet rename

    private async Task Renamesheet(Window mainWindow, string inputExcelCsvFile)
    {
        try
        {   var date = DateTime.Now.ToString().Replace(":", ".");
            using var inputWorkbook = new XLWorkbook(inputExcelCsvFile);
            var readydatasheet = inputWorkbook.Worksheets.FirstOrDefault(ws => ws.Name == "Ready Data");
            var adjusteddatasheet = inputWorkbook.Worksheets.FirstOrDefault(ws => ws.Name == "Adjusted Data");

            if (adjusteddatasheet != null)
            {
                adjusteddatasheet.Name = $"A_{date}";
            }

            if (readydatasheet != null)
            {
                readydatasheet.Name = $"R_{date}";
            }
            
            readydatasheet.SetTabActive(true);
            inputWorkbook.Save();
        }
        catch { }
    }

    #endregion

    #region Run Macro
    private async Task RunMacro(Window mainWindow, string inputExcelCsvFile, string outputCsvFile)
    {
        try
        {
            byte[] decryptedContentBytes = Transformation.GetTransformedFileContent("HSN_Adjustments_GST.dll");
            string vbContent = System.Text.Encoding.UTF8.GetString(decryptedContentBytes);

            string macroRunnerDir = Path.Combine(AppContext.BaseDirectory, "Runner");
            string helperExePath = Path.Combine(macroRunnerDir, "M_HSN_A.exe");

            if (!File.Exists(helperExePath))
            {
                await ShowDialog.ShowMsgBox("Error", $"M_HSN_A.exe not found at: {helperExePath}", "OK", null, 1, mainWindow);
                return;
            }

            //Window_Handler.Minimize(mainWindow);

            var process = new System.Diagnostics.Process();
            process.StartInfo.FileName = helperExePath;
            process.StartInfo.Arguments = $"\"{inputExcelCsvFile}\" \"{outputCsvFile}\" \"HSN_Adjustments_GST\"";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardInput = true;
            process.StartInfo.CreateNoWindow = true;

            process.Start();

            using (var writer = process.StandardInput)
            {
                await writer.WriteAsync(vbContent);
            }

            process.WaitForExit();

            if (process.ExitCode == 0)
            {
               // Window_Handler.Restore(mainWindow);
            }
            else
            {
                await ShowDialog.ShowMsgBox("Error", $"Execution failed. Please review selected files", "OK", null, 1, mainWindow);
            }
        }
        catch (Exception ex)
        {
            await ShowDialog.ShowMsgBox("Error", $"Failed to run macro: {ex.Message}", "OK", null, 1, mainWindow);
        }
    }
    #endregion

    //#region Convert csv to json
    //private async Task ConvertCsvToJson(Window mainWindow, string csvFilePath, string subDir, string HSNType)
    //{
    //    try
    //    {
    //        // Read CSV file
    //        List<hsn_Model_Helpers.HSNEntry> records;
    //        records = hsn_Model_Helpers.CsvManualParser.ParseHSNEntries(csvFilePath);

    //        // Validate data before proceeding
    //        var isNoError = await Check_for_Error(mainWindow, csvFilePath, records);
    //        if (!isNoError)
    //        {
    //            return; // Exit if validation fails
    //        }

    //        // If no errors, proceed with GSTIN and Financial Period input
    //        while (true) // Outer loop for GSTIN and Financial Period input
    //        {
    //            // Show GSTIN Input dialog
    //            var gstinDialog = new ContentDialog
    //            {
    //                Title = "GSTIN Input",
    //                Content = new TextBox { PlaceholderText = "Enter GSTIN" },
    //                PrimaryButtonText = "OK",
    //                CloseButtonText = "Cancel",
    //                XamlRoot = mainWindow.Content.XamlRoot
    //            };

    //            Window_Handler.Restore(mainWindow);
    //            if (await gstinDialog.ShowAsync() != ContentDialogResult.Primary || string.IsNullOrWhiteSpace(((TextBox)gstinDialog.Content).Text))
    //                return;

    //            var gstin = ((TextBox)gstinDialog.Content).Text;

    //            while (true) // Inner loop for Financial Period input
    //            {
    //                var financialPeriodTextBox = new TextBox { PlaceholderText = "Enter MMYYYY", Text = $"{DateTime.Today.Month - 1:D2}{DateTime.Today.Year}" };

    //                // Show Financial Period Input dialog
    //                var fpDialog = new ContentDialog
    //                {
    //                    Title = "Financial Period Input",
    //                    Content = financialPeriodTextBox,
    //                    PrimaryButtonText = "OK",
    //                    SecondaryButtonText = "Back",
    //                    XamlRoot = mainWindow.Content.XamlRoot
    //                };

    //                var fpResult = await fpDialog.ShowAsync();

    //                if (fpResult == ContentDialogResult.Secondary)
    //                    break; // Go back to GSTIN input (outer loop)

    //                if (fpResult != ContentDialogResult.Primary || string.IsNullOrWhiteSpace(financialPeriodTextBox.Text))
    //                    return;

    //                var fp = financialPeriodTextBox.Text;
    //                var outputJsonFile = Path.Combine(subDir, $"{HSNType}_{gstin}_{fp}.json");

    //                // Convert to JSON format using Newtonsoft.Json
    //                var jsonData = new
    //                {
    //                    gstin,
    //                    fp,
    //                    version = "GST3.2.1",
    //                    hash = "hash",
    //                    hsn = new Dictionary<string, object>
    //                    {
    //                        [HSNType] = records.Select((r, i) => new
    //                        {
    //                            num = i + 1,
    //                            hsn_sc = r.HSN,
    //                            desc = r.Description,
    //                            uqc = r.UQC.Split('-')[0],
    //                            qty = Math.Round(r.Total_Quantity, 2),
    //                            rt = r.Rate,
    //                            txval = Math.Round(r.Taxable_Value, 2),
    //                            iamt = Math.Round(r.Integrated_Tax_Amount, 2),
    //                            samt = Math.Round(r.State_UT_Tax_Amount, 2),
    //                            camt = Math.Round(r.Central_Tax_Amount, 2),
    //                            csamt = Math.Round(r.Cess_Amount, 2)
    //                        }).ToList()
    //                    }
    //                };

    //                // Save JSON file using Newtonsoft.Json
    //                await File.WriteAllTextAsync(outputJsonFile, JsonConvert.SerializeObject(jsonData, Formatting.Indented));

    //                var successDialog = new ContentDialog
    //                {
    //                    Title = "Success",
    //                    Content = $"Successfully saved HSN JSON file:\n{outputJsonFile}",
    //                    CloseButtonText = "OK",
    //                    XamlRoot = mainWindow.Content.XamlRoot
    //                };

    //                await successDialog.ShowAsync();

    //                var outputFolderPath = Path.GetDirectoryName(outputJsonFile);
    //                Process.Start("explorer.exe", outputFolderPath);

    //                return; // Exit after successful JSON creation
    //            }
    //        }
    //    }
    //    catch (Exception ex)
    //    {
    //        var errorDialog = new ContentDialog
    //        {
    //            Title = "Error",
    //            Content = $"An error occurred: {ex.Message}",
    //            CloseButtonText = "OK",
    //            XamlRoot = mainWindow.Content.XamlRoot
    //        };
    //        await errorDialog.ShowAsync();
    //    }
    //}

    //string SuggestClosestUQC(string input, IEnumerable<string> validUQCs)
    //{
    //    input = input.Trim().ToUpper();
    //    var inputPrefix = input.Split('-')[0];

    //    foreach (var valid in validUQCs)
    //    {
    //        var validPrefix = valid.Split('-')[0];
    //        if (validPrefix.StartsWith(inputPrefix) || inputPrefix.StartsWith(validPrefix))
    //            return valid;
    //    }

    //    return null;
    //}
    //#endregion

    //#region HSN code Normalization
    //public static string HsnNormalization(string rawHsn, HashSet<string> validHSNs)
    //{
    //    string trimmed = rawHsn.Trim();
    //    if (validHSNs.Contains(trimmed))
    //        return trimmed;

    //    string stripped = trimmed.TrimStart('0');
    //    if (validHSNs.Contains(stripped))
    //        return stripped;

    //    string padded = "0" + stripped;
    //    if (validHSNs.Contains(padded))
    //        return padded;

    //    return null;
    //}
    //#endregion

    //#region Check for error
    //private async Task<bool> Check_for_Error(Window mainWindow, string csvFilePath, List<HSNEntry> records)
    //{
    //    try
    //    {
    //        // Load valid HSN and UQC codes from HSN_Data
    //        var validHSNCodes = HSN_Data.GetHSNCodes();
    //        var validUQCs = HSN_Data.GetUQCCodes();

    //        // Load valid HSN and UQC codes from HSN.json
    //        //var (validHSNCodes, validUQCs) = await hsn_Model_Helpers.LoadValidCodesAsync();
    //        //if (validHSNCodes == null || validUQCs == null)
    //        //{
    //        //    await ShowDialog.ShowMsgBox("Error", "Failed to load HSN.json. Please ensure the file exists in the Misc directory.", "OK", null, 1, mainWindow);
    //        //    return false;
    //        //}

    //        var errors = new List<(int Row, string Message)>();

    //        var combinationRows = new Dictionary<string, List<int>>();
    //        for (int i = 0; i < records.Count; i++)
    //        {
    //            var record = records[i];
    //            var combination = $"{record.HSN?.Trim()}|{record.UQC?.Trim()}|{record.Rate}";
    //            if (!combinationRows.ContainsKey(combination))
    //            {
    //                combinationRows[combination] = new List<int>();
    //            }
    //            combinationRows[combination].Add(i + 2);
    //        }

    //        foreach (var kvp in combinationRows)
    //        {
    //            if (kvp.Value.Count > 1)
    //            {
    //                var combinationParts = kvp.Key.Split('|');
    //                var hsn = combinationParts[0];
    //                var uqc = combinationParts[1];
    //                var rate = combinationParts[2];
    //                foreach (var row in kvp.Value)
    //                {
    //                    errors.Add((row, $"Duplicate entry for HSN: {hsn}, UQC: {uqc}, Rate: {rate}."));
    //                }
    //            }
    //        }

    //        for (int i = 0; i < records.Count; i++)
    //        {
    //            var record = records[i];

    //            if (string.IsNullOrWhiteSpace(record.HSN) || record.HSN == "0")
    //                errors.Add((i + 2, $"HSN is missing or has a value of 0."));

    //            var matchedHSN = HsnNormalization(record.HSN, validHSNCodes);
    //            if (matchedHSN == null)
    //            {
    //                errors.Add((i + 2, $"HSN code '{record.HSN}' not found in master list."));
    //            }
    //            else
    //            {
    //                record.HSN = matchedHSN;
    //            }

    //            if (record.HSN.Length < 2)
    //                errors.Add((i + 2, $"HSN length should be more than three."));

    //            if (string.IsNullOrWhiteSpace(record.UQC) || record.UQC == "0")
    //                errors.Add((i + 2, $"UQC is missing or has a value of 0."));
    //            else if (record.UQC.Trim().ToUpper() == "NA")
    //            {
    //                if (record.Total_Quantity > 0)
    //                    errors.Add((i + 2, $"Total Quantity must be zero when UQC = NA."));
    //            }
    //            else if (!validUQCs.Contains(record.UQC.Trim()))
    //            {
    //                var suggestion = SuggestClosestUQC(record.UQC, validUQCs);
    //                if (suggestion != null)
    //                    errors.Add((i + 2, $"Invalid UQC value '{record.UQC}'. Correct UQC is: {suggestion}"));
    //                else
    //                    errors.Add((i + 2, $"Invalid UQC value '{record.UQC}'."));
    //                if (record.Total_Quantity == 0)
    //                    errors.Add((i + 2, $"Total Quantity cannot be zero."));
    //            }
    //            else
    //            {
    //                if (record.Total_Quantity == 0)
    //                    errors.Add((i + 2, $"Total Quantity cannot be zero."));
    //            }

    //            if (record.Taxable_Value == 0)
    //                errors.Add((i + 2, $"Taxable Value cannot be zero."));

    //            var taxamount = record.Integrated_Tax_Amount + record.Central_Tax_Amount +
    //                record.State_UT_Tax_Amount + record.Cess_Amount;

    //            if (record.Rate > 0 && taxamount == 0)
    //                errors.Add((i + 2, $"Tax amount cannot be zero."));
    //        }

    //        if (errors.Count > 0)
    //        {
    //            var errorMessages = errors.Select(e => $"Row {e.Row}: {e.Message}").ToList();
    //            await ShowDialog.ShowMsgBox1("Data Validation Errors", string.Join(Environment.NewLine, errorMessages), mainWindow);
    //            await HighlightErrorsInReadyDataSheetAsync(csvFilePath, records, errors, mainWindow);
    //            return false;
    //        }

    //        return true;
    //    }
    //    catch (Exception ex)
    //    {
    //        await ShowDialog.ShowMsgBox("Error", $"An error occurred during validation: {ex.Message}", "OK", null, 1, mainWindow);
    //        return false;
    //    }
    //}
    //#endregion

    //#region Highlight error
    //private async Task HighlightErrorsInReadyDataSheetAsync(string csvFilePath, List<HSNEntry> records, List<(int Row, string Message)> errors, Window mainWindow)
    //{
    //    try
    //    {
    //        string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(csvFilePath);
    //        string directory = Path.GetDirectoryName(csvFilePath);
    //        string excelFilePath = Path.Combine(directory, $"{fileNameWithoutExtension}_error_.xlsx");

    //        using (var workbook = new XLWorkbook())
    //        {
    //            var dataWorksheet = workbook.Worksheets.Add("HSN_Highlited");

    //            dataWorksheet.Cell(1, 1).Value = "HSN";
    //            dataWorksheet.Cell(1, 2).Value = "Description";
    //            dataWorksheet.Cell(1, 3).Value = "UQC";
    //            dataWorksheet.Cell(1, 4).Value = "Total Quantity";
    //            dataWorksheet.Cell(1, 5).Value = "Total Value";
    //            dataWorksheet.Cell(1, 6).Value = "Taxable Value";
    //            dataWorksheet.Cell(1, 7).Value = "Integrated Tax Amount";
    //            dataWorksheet.Cell(1, 8).Value = "Central Tax Amount";
    //            dataWorksheet.Cell(1, 9).Value = "State/UT Tax Amount";
    //            dataWorksheet.Cell(1, 10).Value = "Cess Amount";
    //            dataWorksheet.Cell(1, 11).Value = "Rate";

    //            dataWorksheet.Column(4).Style.NumberFormat.Format = "0.00";
    //            dataWorksheet.Column(5).Style.NumberFormat.Format = "0.00";
    //            dataWorksheet.Column(6).Style.NumberFormat.Format = "0.00";
    //            dataWorksheet.Column(7).Style.NumberFormat.Format = "0.00";
    //            dataWorksheet.Column(8).Style.NumberFormat.Format = "0.00";
    //            dataWorksheet.Column(9).Style.NumberFormat.Format = "0.00";
    //            dataWorksheet.Column(10).Style.NumberFormat.Format = "0.00";

    //            for (int i = 0; i < records.Count; i++)
    //            {
    //                var record = records[i];
    //                dataWorksheet.Cell(i + 2, 1).Value = record.HSN;
    //                dataWorksheet.Cell(i + 2, 2).Value = record.Description;
    //                dataWorksheet.Cell(i + 2, 3).Value = record.UQC;
    //                dataWorksheet.Cell(i + 2, 4).Value = record.Total_Quantity;
    //                dataWorksheet.Cell(i + 2, 5).Value = record.Total_Value;
    //                dataWorksheet.Cell(i + 2, 6).Value = record.Taxable_Value;
    //                dataWorksheet.Cell(i + 2, 7).Value = record.Integrated_Tax_Amount;
    //                dataWorksheet.Cell(i + 2, 8).Value = record.Central_Tax_Amount;
    //                dataWorksheet.Cell(i + 2, 9).Value = record.State_UT_Tax_Amount;
    //                dataWorksheet.Cell(i + 2, 10).Value = record.Cess_Amount;
    //                dataWorksheet.Cell(i + 2, 11).Value = record.Rate;
    //            }

    //            foreach (var error in errors)
    //            {
    //                var rowNumber = error.Row;
    //                var row = dataWorksheet.Row(rowNumber);
    //                row.Style.Fill.BackgroundColor = XLColor.Red;
    //            }

    //            var errorWorksheet = workbook.Worksheets.Add("Error Details");

    //            errorWorksheet.Cell(1, 1).Value = "Row";
    //            errorWorksheet.Cell(1, 2).Value = "Error Message";

    //            for (int i = 0; i < errors.Count; i++)
    //            {
    //                errorWorksheet.Cell(i + 2, 1).Value = errors[i].Row;
    //                errorWorksheet.Cell(i + 2, 2).Value = errors[i].Message;
    //            }

    //            dataWorksheet.Columns().AdjustToContents();
    //            errorWorksheet.Columns().AdjustToContents();

    //            await Task.Run(() => workbook.SaveAs(excelFilePath));
    //        }

    //        Process.Start(new ProcessStartInfo(excelFilePath) { UseShellExecute = true });
    //    }
    //    catch (Exception ex)
    //    {
    //        await ShowDialog.ShowMsgBox("Error", $"Failed to convert CSV to Excel, highlight errors, or open the file: {ex.Message}", "OK", null, 1, mainWindow);
    //    }
    //}
    //#endregion

}