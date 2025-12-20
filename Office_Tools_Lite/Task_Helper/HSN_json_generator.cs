using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Serialization;
using ClosedXML.Excel;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using static Office_Tools_Lite.Task_Helper.HSN_json_generator;
using static Office_Tools_Lite.Task_Helper.HSN_CSV_Model_Helpers;

namespace Office_Tools_Lite.Task_Helper;
public static class HSN_json_generator
{
    #region Convert csv to json
    public static async Task ConvertCsvToJson(Window mainWindow, string csvFilePath, string subDir, string HSNType)
    {
        try
        {
            // Read CSV file
            List<HSN_CSV_Model_Helpers.HSN_CSV_Records> records;
            records = HSN_CSV_Model_Helpers.CsvManualParser.ParseHSNEntries(csvFilePath);

            // Validate data before proceeding
            var isNoError = await Check_for_Error(mainWindow, csvFilePath, records);
            if (!isNoError)
            {
                return; // Exit if validation fails
            }

            // If no errors, proceed with GSTIN and Financial Period input
            while (true) // Outer loop for GSTIN and Financial Period input
            {
                var gstinTextBox = new TextBox { PlaceholderText = "Enter GSTIN" };

                // Show GSTIN Input dialog
                var gstinDialog = await ShowDialog.ShowMsgBox(
                    "GSTIN Input", gstinTextBox, "OK", "Cancel", 1,mainWindow);

                Window_Handler.Restore(mainWindow);
                if (gstinDialog != ContentDialogResult.Primary ||
                    string.IsNullOrWhiteSpace(gstinTextBox.Text))
                {
                    return;
                }

                var gstin = gstinTextBox.Text.Trim();

                while (true) // Inner loop for Financial Period input
                {
                    var financialPeriodTextBox = new TextBox { PlaceholderText = "Enter MMYYYY", Text = $"{DateTime.Today.Month - 1:D2}{DateTime.Today.Year}" };

                    // Show Financial Period Input dialog
                    var fpDialog = await ShowDialog.ShowMsgBox(
                        "Financial Period Input",
                        financialPeriodTextBox, "OK", "Back", 1, mainWindow);

                    if (fpDialog == ContentDialogResult.Secondary)
                        break; // Back to GSTIN

                    if (fpDialog != ContentDialogResult.Primary ||
                        string.IsNullOrWhiteSpace(financialPeriodTextBox.Text))
                    {
                        return;
                    }

                    var fp = financialPeriodTextBox.Text;
                    var outputJsonFile = Path.Combine(subDir, $"{HSNType}_{gstin}_{fp}.json");

                    var jsonData = new HSNJsonData
                    {
                        gstin = gstin,
                        fp = fp,
                        version = "GST3.2.1",
                        hash = "hash",
                        hsn = new Dictionary<string, List<HSNRecord>>
                        {
                            [HSNType] = records.Select((r, i) => new HSNRecord
                            {
                                num = i + 1,
                                hsn_sc = r.HSN,
                                desc = r.Description,
                                uqc = r.UQC.Split('-')[0],
                                qty = Math.Round(r.Total_Quantity, 2),
                                rt = r.Rate,
                                txval = Math.Round(r.Taxable_Value, 2),
                                iamt = Math.Round(r.Integrated_Tax_Amount, 2),
                                samt = Math.Round(r.State_UT_Tax_Amount, 2),
                                camt = Math.Round(r.Central_Tax_Amount, 2),
                                csamt = Math.Round(r.Cess_Amount, 2)
                            }).ToList()
                        }
                    };

                    await File.WriteAllTextAsync(outputJsonFile,
                            JsonSerializer.Serialize(jsonData, HSNOutputJsonContext.Default.HSNJsonData));

                    await ShowDialog.ShowMsgBox("Success", $"Successfully saved HSN JSON file:\n{outputJsonFile}", "OK", null, 1, mainWindow);

                    var outputFolderPath = Path.GetDirectoryName(outputJsonFile);
                    Process.Start("explorer.exe", outputFolderPath);

                    return; // Exit after successful JSON creation
                }
            }
        }
        catch (Exception ex)
        {
            await ShowDialog.ShowMsgBox("Error", $"An error occurred: {ex.Message}", "OK", null, 1, mainWindow);
            
        }
    }

    static string SuggestClosestUQC(string input, IEnumerable<string> validUQCs)
    {
        input = input.Trim().ToUpper();
        var inputPrefix = input.Split('-')[0];

        foreach (var valid in validUQCs)
        {
            var validPrefix = valid.Split('-')[0];
            if (validPrefix.StartsWith(inputPrefix) || inputPrefix.StartsWith(validPrefix))
                return valid;
        }

        return null;
    }
    #endregion

    #region HSN code Normalization
    public static string HsnNormalization(string rawHsn, HashSet<string> validHSNs)
    {
        string trimmed = rawHsn.Trim();
        if (validHSNs.Contains(trimmed))
            return trimmed;

        string stripped = trimmed.TrimStart('0');
        if (validHSNs.Contains(stripped))
            return stripped;

        string padded = "0" + stripped;
        if (validHSNs.Contains(padded))
            return padded;

        return null;
    }
    #endregion

    #region Check for error
    public static async Task<bool> Check_for_Error(Window mainWindow, string csvFilePath, List<HSN_CSV_Records> records)
    {
        try
        {
            // Load valid HSN and UQC codes from HSN_Data
            //var validHSNCodes = HSN_Data.GetHSNCodes();
            //var validUQCs = HSN_Data.GetUQCCodes();
            //var validHSNCodes = HSN_Data.GetHSNCodes();     // HashSet<string>
            //var validUQCs = HSN_Data.GetUQCCodes();      // HashSet<string>

            // Load valid HSN and UQC codes from HSN.json
            var (validHSNCodes, descMap, validUQCs) = await HSN_CSV_Model_Helpers.LoadValidCodesAsync();

            /// Example for looking HSN Description by code
            // Lookup description for HSN "01"
            //string hsnToCheck = "01";
            //if (descMap.TryGetValue(hsnToCheck, out string descr))
            //{
            //    await ShowDialog.ShowMsgBox1("HSN", $"HSN {hsnToCheck}: {descr}", mainWindow);
            //}
            //else
            //{
            //    await ShowDialog.ShowMsgBox1("HSN", $"HSN {hsnToCheck} not found.", mainWindow);
            //}


            if (validHSNCodes == null || validUQCs == null)
            {
                await ShowDialog.ShowMsgBox("Error", "Failed to load HSN.json.", "OK", null, 1, mainWindow);
                return false;
            }

            var errors = new List<(int Row, string Message)>();

            var combinationRows = new Dictionary<string, List<int>>();
            for (int i = 0; i < records.Count; i++)
            {
                var record = records[i];
                var combination = $"{record.HSN?.Trim()}|{record.UQC?.Trim()}|{record.Rate}";
                if (!combinationRows.ContainsKey(combination))
                {
                    combinationRows[combination] = new List<int>();
                }
                combinationRows[combination].Add(i + 2);
            }

            foreach (var kvp in combinationRows)
            {
                if (kvp.Value.Count > 1)
                {
                    var combinationParts = kvp.Key.Split('|');
                    var hsn = combinationParts[0];
                    var uqc = combinationParts[1];
                    var rate = combinationParts[2];
                    foreach (var row in kvp.Value)
                    {
                        errors.Add((row, $"Duplicate entry for HSN: {hsn}, UQC: {uqc}, Rate: {rate}."));
                    }
                }
            }

            for (int i = 0; i < records.Count; i++)
            {
                var record = records[i];

                if (string.IsNullOrWhiteSpace(record.HSN) || record.HSN == "0")
                    errors.Add((i + 2, $"HSN is missing or has a value of 0."));

                var matchedHSN = HsnNormalization(record.HSN, validHSNCodes);
                if (matchedHSN == null)
                {
                    errors.Add((i + 2, $"HSN code '{record.HSN}' not found in master list."));
                }
                else
                {
                    record.HSN = matchedHSN;
                }

                if (record.HSN.Length < 2)
                    errors.Add((i + 2, $"HSN length should be more than three."));

                if (string.IsNullOrWhiteSpace(record.UQC) || record.UQC == "0")
                    errors.Add((i + 2, $"UQC is missing or has a value of 0."));
                else if (record.UQC.Trim().ToUpper() == "NA")
                {
                    if (record.Total_Quantity > 0)
                        errors.Add((i + 2, $"Total Quantity must be zero when UQC = NA."));
                }
                else if (!validUQCs.Contains(record.UQC.Trim()))
                {
                    var suggestion = SuggestClosestUQC(record.UQC, validUQCs);
                    if (suggestion != null)
                        errors.Add((i + 2, $"Invalid UQC value '{record.UQC}'. Correct UQC is: {suggestion}"));
                    else
                        errors.Add((i + 2, $"Invalid UQC value '{record.UQC}'."));
                    if (record.Total_Quantity == 0)
                        errors.Add((i + 2, $"Total Quantity cannot be zero."));
                }
                else
                {
                    if (record.Total_Quantity == 0)
                        errors.Add((i + 2, $"Total Quantity cannot be zero."));
                }

                if (record.Taxable_Value == 0)
                    errors.Add((i + 2, $"Taxable Value cannot be zero."));

                var taxamount = record.Integrated_Tax_Amount + record.Central_Tax_Amount +
                    record.State_UT_Tax_Amount + record.Cess_Amount;

                if (record.Rate > 0 && taxamount == 0)
                    errors.Add((i + 2, $"Tax amount cannot be zero."));
            }

            if (errors.Count > 0)
            {
                var errorMessages = errors.Select(e => $"Row {e.Row}: {e.Message}").ToList();
                await ShowDialog.ShowMsgBox("Data Validation Errors", string.Join(Environment.NewLine, errorMessages), "Ok", null, 1, mainWindow);
                await HighlightErrorsInReadyDataSheetAsync(csvFilePath, records, errors, mainWindow);
                return false;
            }

            return true;
        }
        catch (Exception ex)
        {
            await ShowDialog.ShowMsgBox("Error", $"An error occurred during validation: {ex.Message}", "Ok", null, 1, mainWindow);
            return false;
        }
    }
    #endregion

    #region Highlight error
    private static async Task HighlightErrorsInReadyDataSheetAsync(string csvFilePath, List<HSN_CSV_Records> records, List<(int Row, string Message)> errors, Window mainWindow)
    {
        try
        {
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(csvFilePath);
            string directory = Path.GetDirectoryName(csvFilePath);
            string excelFilePath = Path.Combine(directory, $"{fileNameWithoutExtension}_error_.xlsx");

            using (var workbook = new XLWorkbook())
            {
                var dataWorksheet = workbook.Worksheets.Add("HSN_Highlited");

                dataWorksheet.Cell(1, 1).Value = "HSN";
                dataWorksheet.Cell(1, 2).Value = "Description";
                dataWorksheet.Cell(1, 3).Value = "UQC";
                dataWorksheet.Cell(1, 4).Value = "Total Quantity";
                dataWorksheet.Cell(1, 5).Value = "Total Value";
                dataWorksheet.Cell(1, 6).Value = "Taxable Value";
                dataWorksheet.Cell(1, 7).Value = "Integrated Tax Amount";
                dataWorksheet.Cell(1, 8).Value = "Central Tax Amount";
                dataWorksheet.Cell(1, 9).Value = "State/UT Tax Amount";
                dataWorksheet.Cell(1, 10).Value = "Cess Amount";
                dataWorksheet.Cell(1, 11).Value = "Rate";

                dataWorksheet.Column(4).Style.NumberFormat.Format = "0.00";
                dataWorksheet.Column(5).Style.NumberFormat.Format = "0.00";
                dataWorksheet.Column(6).Style.NumberFormat.Format = "0.00";
                dataWorksheet.Column(7).Style.NumberFormat.Format = "0.00";
                dataWorksheet.Column(8).Style.NumberFormat.Format = "0.00";
                dataWorksheet.Column(9).Style.NumberFormat.Format = "0.00";
                dataWorksheet.Column(10).Style.NumberFormat.Format = "0.00";

                for (int i = 0; i < records.Count; i++)
                {
                    var record = records[i];
                    dataWorksheet.Cell(i + 2, 1).Value = record.HSN;
                    dataWorksheet.Cell(i + 2, 2).Value = record.Description;
                    dataWorksheet.Cell(i + 2, 3).Value = record.UQC;
                    dataWorksheet.Cell(i + 2, 4).Value = record.Total_Quantity;
                    dataWorksheet.Cell(i + 2, 5).Value = record.Total_Value;
                    dataWorksheet.Cell(i + 2, 6).Value = record.Taxable_Value;
                    dataWorksheet.Cell(i + 2, 7).Value = record.Integrated_Tax_Amount;
                    dataWorksheet.Cell(i + 2, 8).Value = record.Central_Tax_Amount;
                    dataWorksheet.Cell(i + 2, 9).Value = record.State_UT_Tax_Amount;
                    dataWorksheet.Cell(i + 2, 10).Value = record.Cess_Amount;
                    dataWorksheet.Cell(i + 2, 11).Value = record.Rate;
                }

                foreach (var error in errors)
                {
                    var rowNumber = error.Row;
                    var row = dataWorksheet.Row(rowNumber);
                    row.Style.Fill.BackgroundColor = XLColor.Red;
                }

                var errorWorksheet = workbook.Worksheets.Add("Error Details");

                errorWorksheet.Cell(1, 1).Value = "Row";
                errorWorksheet.Cell(1, 2).Value = "Error Message";

                for (int i = 0; i < errors.Count; i++)
                {
                    errorWorksheet.Cell(i + 2, 1).Value = errors[i].Row;
                    errorWorksheet.Cell(i + 2, 2).Value = errors[i].Message;
                }

                dataWorksheet.Columns().AdjustToContents();
                errorWorksheet.Columns().AdjustToContents();

                await Task.Run(() => workbook.SaveAs(excelFilePath));
            }

            Process.Start(new ProcessStartInfo(excelFilePath) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            await ShowDialog.ShowMsgBox("Error", $"Failed to convert CSV to Excel, highlight errors, or open the file: {ex.Message}", "Ok", null, 1, mainWindow);
        }
    }
    #endregion

    public class HSNJsonData
    {
        public string gstin { get; set; }
        public string fp { get; set; }
        public string version { get; set; }
        public string hash { get; set; }
        public Dictionary<string, List<HSNRecord>> hsn { get; set; }
    }

    public class HSNRecord
    {
        public int num { get; set; }
        public string hsn_sc { get; set; }
        public string desc { get; set; }
        public string uqc { get; set; }
        public decimal qty { get; set; }
        public int rt { get; set; }
        public decimal txval { get; set; }
        public decimal iamt { get; set; }
        public decimal samt { get; set; }
        public decimal camt { get; set; }
        public decimal csamt { get; set; }
    }
                                
}
[JsonSourceGenerationOptions(WriteIndented = true)]
[JsonSerializable(typeof(HSNJsonData))]
[JsonSerializable(typeof(Dictionary<string, List<HSNRecord>>))]
[JsonSerializable(typeof(List<HSNRecord>))]
[JsonSerializable(typeof(HSNRecord))]
public partial class HSNOutputJsonContext : JsonSerializerContext
{
}
