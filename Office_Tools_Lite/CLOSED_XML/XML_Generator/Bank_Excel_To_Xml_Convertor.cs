using System.Data;
using System.Diagnostics;
using System.Xml.Linq;
using ClosedXML.Excel;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Office_Tools_Lite.Task_Helper;
using Windows.ApplicationModel.DataTransfer;
using Windows.Storage.Pickers;
using WinRT.Interop;

namespace Office_Tools_Lite.CLOSED_XML.XML_Generator;

public class Bank_Excel_To_Xml_Converter
{
    private DataTable _table;

    #region Execute Main task
    public async Task Execute(Window mainWindow)
    {
        // File picker dialog
        var filePicker = new FileOpenPicker();
        var hwnd = WindowNative.GetWindowHandle(mainWindow);
        InitializeWithWindow.Initialize(filePicker, hwnd);

        filePicker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
        filePicker.FileTypeFilter.Add(".xlsx");

        var file = await filePicker.PickSingleFileAsync();
        if (file != null)
        {
            string FinancialYear = await ShowInputDialog("Enter FY (YYYY-YY):", mainWindow);
            if (string.IsNullOrEmpty(FinancialYear))
            {
                return; // User canceled, return safely
            }

            // Validate Financial Year format and extract years
            if (!ValidateFinancialYear(FinancialYear, out int year1, out int year2))
            {
                await ShowDialog.ShowMsgBox("Financial Year Error", "Please review Financial Year", "OK", null, 1, mainWindow);
                return;
            }

            try
            {
                string excelPath = file.Path; // Use file.Path instead of dialog.FileName
                _table = await ReadExcel(excelPath);
                if (_table == null) return;
                List<int> NoParticulars;
                List<int> noValue;
                List<int> NoTransTypes;

                // Validate dates in the Excel data
                if (!ValidateExcelData(_table, year1, year2, out List<(int Row, string Date)> emptyDates, out List<(int Row, string Date)> invalidDateFormats, out List<(int Row, string Date)> incorrectYearDates, out NoParticulars, out noValue, out NoTransTypes))
                {
                    await ShowDialog.ShowMsgBox("Date Validation Error", "There are few errors in Excel data...!", "OK", null, 1, mainWindow);
                    await HighlightErrorsInReadyDataSheetAsync(excelPath, emptyDates, invalidDateFormats, incorrectYearDates, NoParticulars, noValue, NoTransTypes);
                    return;
                }

                XDocument xml = BuildBankVoucherXml(_table);
                string outputPath = Path.Combine(Path.GetDirectoryName(excelPath), "Bank_Vouchers.xml");
                xml.Save(outputPath);
                await ShowDialog.ShowMsgBox("Success", "XML file generated:\n" + outputPath, "Ok", null, 1, mainWindow);

                // Open the output folder in Explorer
                var outputFolderPath = Path.GetDirectoryName(outputPath); // Get the directory of the output file
                System.Diagnostics.Process.Start("explorer.exe", outputFolderPath);
                var dataPackage = new DataPackage();
                dataPackage.SetText(outputPath);
                Clipboard.SetContent(dataPackage);

            }
            catch (Exception ex)
            {
                await ShowDialog.ShowMsgBox("Exception", $"Error: {ex.Message}", "OK", null, 1, mainWindow);
            }
        }
    }
    #endregion

    #region Read Excel Data
    private async Task<DataTable> ReadExcel(string path)
    {
        try
        {
            using (var workbook = new XLWorkbook(path))
            {
                var sheetName = "Ready Data";
                if (!workbook.Worksheets.Contains(sheetName))
                {
                    await ShowDialog.ShowMsgBox("Error", "Worksheet 'Ready Data' not found in the Excel file.", "OK", null, 1, App.MainWindow);
                    return null;
                }

                var worksheet = workbook.Worksheet(sheetName);
                // Disable AutoFilter if it exists
                if (worksheet.AutoFilter.IsEnabled)
                {
                    worksheet.AutoFilter.Clear();
                }

                var range = worksheet.RangeUsed();
                if (range == null || range.RowCount() == 0)
                {
                    await ShowDialog.ShowMsgBox("Error", "No data found in the 'Ready Data' worksheet.", "OK", null, 1, App.MainWindow);
                    return null;
                }

                return range.AsTable().AsNativeDataTable();
            }
        }
        catch (Exception ex)
        {
            await ShowDialog.ShowMsgBox("Error", $"Error reading Excel file: {ex.Message}", "OK", null, 1, App.MainWindow);
            return null;
        }
    }
    #endregion

    #region Build XML
    private XDocument BuildBankVoucherXml(DataTable table)
    {
        //var vouchers = table.AsEnumerable().Select(row => CreateBankVoucherElement(row, table));
        int maxRows = 15; // LIMIT to 15 rows
        var vouchers = table.AsEnumerable().Take(maxRows).Select(row => CreateBankVoucherElement(row, table));

        return new XDocument(
            new XElement("ENVELOPE",
                new XElement("HEADER",
                    new XElement("TALLYREQUEST", "Import Data")
                ),
                new XElement("BODY",
                    new XElement("IMPORTDATA",
                        new XElement("REQUESTDESC",
                            new XElement("REPORTNAME", "All Masters"),
                            new XElement("STATICVARIABLES",
                                new XElement("SVCURRENTCOMPANY", " ")
                            )
                        ),
                        new XElement("REQUESTDATA", vouchers)
                    )
                )
            )
        );
    }
    #endregion

    #region Create Voucher Element
    private XElement CreateBankVoucherElement(DataRow row, DataTable table)
    {
        string date = Convert.ToDateTime(row["Date"]).ToString("yyyyMMdd");
        string TransType = row["Transaction Type"].ToString().Trim();
        string vchType = table.Columns.Contains("Voucher Type") && !string.IsNullOrWhiteSpace(row["Voucher Type"]?.ToString()) ? row["Voucher Type"].ToString() : TransType;
        string bankname = row["Bank Name"].ToString();
        string particulars = row["Particulars"].ToString();

        // Handle Narration & Voucher No. as optional
        string narration = table.Columns.Contains("Narration") ? row["Narration"].ToString() : "";
        string voucherNo = table.Columns.Contains("Voucher No.") ? row["Voucher No."].ToString() : "";

        decimal amount = 0;
        bool isDebit = false;

        if (decimal.TryParse(row["Debit"]?.ToString(), out decimal debitAmount) && debitAmount > 0)
        {
            amount = debitAmount;
            isDebit = true;
        }
        else if (decimal.TryParse(row["Credit"]?.ToString(), out decimal creditAmount) && creditAmount > 0)
        {
            amount = creditAmount;
            isDebit = false;
        }
        else
        {
            throw new ArgumentException("Both Debit and Credit are empty or invalid.");
        }

        XElement voucherElement = new XElement("VOUCHER",
            new XAttribute("VCHTYPE", vchType),
            new XAttribute("ACTION", "Create"),
            new XAttribute("OBJVIEW", "Accounting Voucher View"),
            new XElement("DATE", date),
            new XElement("VOUCHERTYPENAME", vchType),
            new XElement("PARTYLEDGERNAME", bankname),
            new XElement("ISINVOICE", "No"),
            new XElement("EFFECTIVEDATE", date)
        );

        // Add NARRATION and VOUCHERNUMBER directly to voucherElement
        if (!string.IsNullOrEmpty(narration))
        {
            voucherElement.Add(new XElement("NARRATION", narration));
        }
        if (!string.IsNullOrEmpty(voucherNo))
        {
            voucherElement.Add(new XElement("VOUCHERNUMBER", voucherNo));
        }

        // Build debit and credit entries
        XElement debitEntry = new XElement("LEDGERENTRIES.LIST",
            new XElement("LEDGERNAME", isDebit ? bankname : particulars),
            new XElement("ISDEEMEDPOSITIVE", "No"),
            new XElement("ISPARTYLEDGER", "Yes"),
            new XElement("AMOUNT", amount.ToString("F2"))
        );

        XElement creditEntry = new XElement("LEDGERENTRIES.LIST",
            new XElement("LEDGERNAME", isDebit ? particulars : bankname),
            new XElement("ISDEEMEDPOSITIVE", "Yes"),
            new XElement("ISPARTYLEDGER", "Yes"),
            new XElement("AMOUNT", (-amount).ToString("F2"))
        );

        // Correct the order for display in Tally
        if (TransType.Equals("Receipt", StringComparison.OrdinalIgnoreCase))
        {
            // Debit (To) first, then Credit (By)
            voucherElement.Add(debitEntry);
            voucherElement.Add(creditEntry);
        }
        else if (TransType.Equals("Payment", StringComparison.OrdinalIgnoreCase))
        {
            // Credit (By) first, then Debit (To)
            voucherElement.Add(creditEntry);
            voucherElement.Add(debitEntry);
        }
        else if (TransType.Equals("Journal", StringComparison.OrdinalIgnoreCase))
        {
            // Credit (By) first, then Debit (To)
            voucherElement.Add(creditEntry);
            voucherElement.Add(debitEntry);
        }
        else
        {
            // Default order: Debit first
            voucherElement.Add(debitEntry);
            voucherElement.Add(creditEntry);
        }

        return new XElement("TALLYMESSAGE", voucherElement);
    }
    #endregion

    #region Validate Financial year
    private bool ValidateFinancialYear(string financialYear, out int year1, out int year2)
    {
        year1 = 0;
        year2 = 0;
        string errorMessage = string.Empty;

        // Check if the financial year matches the format YYYY-YY
        if (!System.Text.RegularExpressions.Regex.IsMatch(financialYear, @"^\d{4}-\d{2}$"))
        {
            errorMessage = "Financial year must be in the format YYYY-YY (e.g., 2024-25).";
            return false;
        }

        // Split the financial year
        var parts = financialYear.Split('-');
        if (!int.TryParse(parts[0], out year1) || !int.TryParse(parts[1], out int shortYear))
        {
            errorMessage = "Invalid year format in financial year input.";
            return false;
        }

        // Convert short year to full year
        year2 = 2000 + shortYear;

        // Validate that year2 is year1 + 1
        if (year2 != year1 + 1)
        {
            errorMessage = $"Invalid financial year: {year2} should be {year1 + 1} (e.g., for 2024-25, expect 2025 as second year).";
            return false;
        }

        // Ensure years are reasonable (e.g., not in the future beyond current year or too far in the past)
        int currentYear = DateTime.Now.Year;
        if (year1 < 1900 || year1 > currentYear)
        {
            errorMessage = $"Year {year1} is out of valid range (1900-{currentYear}).";
            return false;
        }

        return true;
    }
    #endregion

    #region Validate Excel Data
    private bool ValidateExcelData(DataTable table, int year1, int year2, out List<(int Row, string Date)> emptyDates, out List<(int Row, string Date)> invalidDateFormats, out List<(int Row, string Date)> incorrectYearDates, out List<int> NoParticulars, out List<int> Novalue, out List<int> noTransTypes)
    {
        string errorMessage = string.Empty;
        emptyDates = new List<(int Row, string Date)>();
        invalidDateFormats = new List<(int Row, string Date)>();
        incorrectYearDates = new List<(int Row, string Date)>();
        NoParticulars = new List<int>();
        Novalue = new List<int>();
        noTransTypes = new List<int>();

        string[] valCols = { "Debit", "Credit" };

        if (!table.Columns.Contains("Date"))
        {
            errorMessage = "Excel file must contain a 'Date' column.";
            return false;
        }

        int rowIndex = 1; // Excel row index for user readability

        foreach (DataRow row in table.Rows)
        {
            rowIndex++;
            string dateStr = row["Date"]?.ToString();
            if (string.IsNullOrEmpty(dateStr))
            {
                emptyDates.Add((rowIndex, "Empty"));
                continue;
            }

            DateTime date;
            bool parsed = false;

            // If already DateTime
            if (row["Date"] is DateTime excelDate)
            {
                date = excelDate;
                parsed = true;
                dateStr = date.ToString("dd-MM-yyyy");
            }
            else
            {
                string[] formats = { "dd/MM/yy", "dd-MM-yy", "dd/MM/yyyy", "dd-MM-yyyy" };
                parsed = DateTime.TryParseExact(dateStr, formats,
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out date);
                if (parsed)
                {
                    dateStr = date.ToString("dd-MM-yyyy");
                }
            }

            if (!parsed)
            {
                invalidDateFormats.Add((rowIndex, dateStr));
                continue;
            }

            // Normalize year for 2-digit years
            int year = date.Year;
            if (year < 100)
            {
                if (year >= (year1 % 100) && year <= (year2 % 100))
                {
                    year = year <= (year1 % 100) + 1 ? year1 : year2;
                }
                else
                {
                    year += 2000;
                }
            }

            // Financial year check
            int month = date.Month;
            if (month >= 4 && month <= 12)
            {
                if (year != year1)
                    incorrectYearDates.Add((rowIndex, dateStr));
            }
            else if (month >= 1 && month <= 3)
            {
                if (year != year2)
                    incorrectYearDates.Add((rowIndex, dateStr));
            }

            // Bank Name + Particulars check
            string particulars = table.Columns.Contains("Particulars") ? row["Particulars"].ToString() : string.Empty;
            string bankName = table.Columns.Contains("Bank Name") ? row["Bank Name"].ToString() : string.Empty;

            decimal particularsDec;
            bool isParticularsEmpty = string.IsNullOrWhiteSpace(particulars) ||
                                      (decimal.TryParse(particulars, out particularsDec) && particularsDec == 0);

            bool isBankNameEmpty = string.IsNullOrWhiteSpace(bankName);

            if (isParticularsEmpty || isBankNameEmpty)
            {
                NoParticulars.Add(rowIndex);
            }


            // Value check (Debit + Credit)
            decimal valTotal = 0;
            foreach (string col in valCols)
            {
                if (table.Columns.Contains(col))
                {
                    decimal val;
                    if (decimal.TryParse(row[col].ToString(), out val))
                        valTotal += val;
                }
            }
            if (valTotal == 0)
                Novalue.Add(rowIndex);

            // ===== NoVoucherType Check =====
            string voucherType = table.Columns.Contains("Transaction Type") ? row["Transaction Type"].ToString() : string.Empty;
            if (string.IsNullOrWhiteSpace(voucherType))
            {
                noTransTypes.Add(rowIndex);
            }
        }

        // If any errors found → false
        if (emptyDates.Any() || invalidDateFormats.Any() || incorrectYearDates.Any() || NoParticulars.Any() || Novalue.Any() || noTransTypes.Any())
        {
            return false;
        }

        return true;
    }
    #endregion

    #region Highlight Errors
    private async Task HighlightErrorsInReadyDataSheetAsync(string excelPath, List<(int Row, string Date)> emptyDates, List<(int Row, string Date)> invalidDateFormats, List<(int Row, string Date)> incorrectYearDates,
        List<int> noParticulars, List<int> noValues, List<int> noTransTypes)
    {
        try
        {
            // Ensure the file exists
            if (!File.Exists(excelPath))
            {
                await ShowDialog.ShowMsgBox("Error", "Excel file does not exist.", "OK", null, 1, App.MainWindow);
                return;
            }

            // Open the file and highlight errors in the "Ready Data" sheet
            using (var workbook = new XLWorkbook(excelPath))
            {
                var worksheet = workbook.Worksheet("Ready Data");
                if (worksheet == null)
                {
                    await ShowDialog.ShowMsgBox("Error", "'Ready Data' sheet not found in the file.", "OK", null, 1, App.MainWindow);
                    return;
                }

                // Determine the last column with data
                int lastColumn = worksheet.LastColumnUsed().ColumnNumber();

                // Highlight rows with errors
                foreach (var error in emptyDates)
                {
                    int row = error.Row;
                    var rowRange = worksheet.Range(row, 1, row, lastColumn);
                    rowRange.Style.Fill.BackgroundColor = XLColor.Yellow;
                }

                foreach (var error in invalidDateFormats)
                {
                    int row = error.Row;
                    var rowRange = worksheet.Range(row, 1, row, lastColumn);
                    rowRange.Style.Fill.BackgroundColor = XLColor.Blue;
                }

                foreach (var error in incorrectYearDates)
                {
                    int row = error.Row;
                    var rowRange = worksheet.Range(row, 1, row, lastColumn);
                    rowRange.Style.Fill.BackgroundColor = XLColor.Red;
                }

                foreach (var row in noParticulars)
                {
                    var rowRange = worksheet.Range(row, 1, row, lastColumn);
                    rowRange.Style.Fill.BackgroundColor = XLColor.Orange;
                }

                foreach (var row in noValues)
                {
                    var rowRange = worksheet.Range(row, 1, row, lastColumn);
                    rowRange.Style.Fill.BackgroundColor = XLColor.Red;
                }
                // NoVoucherType
                foreach (var row in noTransTypes)
                {
                    var rowRange = worksheet.Range(row, 1, row, lastColumn);
                    rowRange.Style.Fill.BackgroundColor = XLColor.Pink;
                }

                // Add a new sheet for error details
                var errorSheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == "Error Details") ?? workbook.Worksheets.Add("Error Details");

                // Clear existing content if the sheet already exists
                errorSheet.Clear();

                // Write the log data into the sheet
                errorSheet.Cell(1, 1).Value = "Bank Voucher Error Details";
                errorSheet.Cell(2, 1).Value = $"Generated On: {DateTime.Now:dd-MM-yyyy hh:mm:ss tt}";

                // Write Empty Dates details horizontally
                errorSheet.Cell(4, 1).Value = "Empty Dates";
                int colIndex = 2;
                foreach (var error in emptyDates)
                {
                    errorSheet.Cell(4, colIndex++).Value = $"Row {error.Row}";
                }

                // Write Invalid Date Formats details horizontally
                errorSheet.Cell(5, 1).Value = "Invalid Date Formats";
                colIndex = 2;
                foreach (var error in invalidDateFormats)
                {
                    errorSheet.Cell(5, colIndex++).Value = $"Row {error.Row}: {error.Date}";
                }

                // Write Incorrect Year Dates details horizontally
                errorSheet.Cell(6, 1).Value = "Incorrect Year Dates";
                colIndex = 2;
                foreach (var error in incorrectYearDates)
                {
                    errorSheet.Cell(6, colIndex++).Value = $"Row {error.Row}: {error.Date}";
                }

                errorSheet.Cell(7, 1).Value = "No Particulars/Bank Name";
                colIndex = 2;
                foreach (var row in noParticulars)
                {
                    errorSheet.Cell(7, colIndex++).Value = $"Row {row}";
                }

                errorSheet.Cell(8, 1).Value = "No Values";
                colIndex = 2;
                foreach (var row in noValues)
                {
                    errorSheet.Cell(8, colIndex++).Value = $"Row {row}";
                }

                errorSheet.Cell(9, 1).Value = "No Transaction Type";
                colIndex = 2;
                foreach (var row in noTransTypes)
                {
                    errorSheet.Cell(9, colIndex++).Value = $"Row {row}";
                }

                // Auto-fit columns for better readability
                errorSheet.Columns().AdjustToContents();

                // Mark the "Error Details" sheet as active
                errorSheet.SetTabActive(false);

                if (worksheet.AutoFilter.IsEnabled)
                {
                    worksheet.AutoFilter.Clear();
                }
                // Save changes to the file
                workbook.Save();
            }

            // Try to open the file after saving
            try
            {
                Process.Start(new ProcessStartInfo(excelPath) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                await ShowDialog.ShowMsgBox("Error", $"Failed to open the Excel file: {ex.Message}", "OK", null, 1, App.MainWindow);
            }
        }
        catch (Exception ex)
        {
            await ShowDialog.ShowMsgBox("Error", $"An error occurred while updating the Excel file: {ex.Message}", "OK", null, 1, App.MainWindow);
        }
    }
    #endregion

    #region Show Input Dialog
    private async Task<string> ShowInputDialog(string title, Window mainWindow)
    {
        var inputTextBox = new TextBox
        {
            AcceptsReturn = false,
            Height = 32
        };

        var dialog = await ShowDialog.ShowMsgBox(title, inputTextBox, "OK", "Cancel", 1, mainWindow);

        if (dialog == ContentDialogResult.Primary)
        {
            return inputTextBox.Text;
        }
        else
        {
            return null;
        }
    }
    #endregion

}