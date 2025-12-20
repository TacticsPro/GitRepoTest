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

public class Voucher_Excel_To_Xml_Converter
{
    private DataTable _table;

    #region Execute Main Task
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

                List<int> noParticulars;
                List<int> noValues;
                List<int> noTransTypes;

                // Validate dates in the Excel data
                if (!ValidateExcelData(_table, year1, year2, out List<(int Row, string Date)> emptyDates, out List<(int Row, string Date)> invalidDateFormats, out List<(int Row, string Date)> incorrectYearDates, out noParticulars, out noValues, out noTransTypes))
                {
                    await ShowDialog.ShowMsgBox("Date Validation Error", "There are few errors in Excel data...!", "OK", null, 1, mainWindow);
                    await HighlightErrorsInReadyDataSheetAsync(excelPath, emptyDates, invalidDateFormats, incorrectYearDates, noParticulars, noValues, noTransTypes);
                    return;
                }

                XDocument xml = BuildVoucherXml(_table);
                string outputPath = Path.Combine(Path.GetDirectoryName(excelPath), "Vouchers.xml");
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

    #region Read Excle Data
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
    private XDocument BuildVoucherXml(DataTable table)
    {
        //var vouchers = table.AsEnumerable().Skip(1).Select(row => CreateBankVoucherElement(row, table));
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
                                new XElement("SVCURRENTCOMPANY", "Sample Data")
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
        string ledger1 = row["Ledger 1"].ToString();
        string ledger2 = row["Ledger 2"].ToString();
        string amountStr = row["Amount"].ToString();
        decimal amount = Convert.ToDecimal(amountStr);
        // Handle Narration & Voucher No. as optional
        string narration = table.Columns.Contains("Narration") ? row["Narration"].ToString() : "";
        string voucherNo = table.Columns.Contains("Voucher No.") ? row["Voucher No."].ToString() : "";

        XElement voucherElement = new XElement("VOUCHER",
            new XAttribute("VCHTYPE", vchType),
            new XAttribute("ACTION", "Create"),
            new XAttribute("OBJVIEW", "Accounting Voucher View"),
            new XElement("DATE", date),
            new XElement("VOUCHERTYPENAME", vchType),
            new XElement("PARTYLEDGERNAME", ledger2),
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

        // Build debit and credit entries for Payment and Journal
        XElement debitEntry_Payment_Journal = new XElement("LEDGERENTRIES.LIST",
            new XElement("LEDGERNAME", ledger2),
            new XElement("ISDEEMEDPOSITIVE", "No"),
            new XElement("ISPARTYLEDGER", "Yes"),
            new XElement("AMOUNT", amount.ToString("F2"))
        );

        XElement creditEntry_Payment_Journal = new XElement("LEDGERENTRIES.LIST",
            new XElement("LEDGERNAME", ledger1),
            new XElement("ISDEEMEDPOSITIVE", "Yes"),
            new XElement("ISPARTYLEDGER", "Yes"),
            new XElement("AMOUNT", (-amount).ToString("F2"))
        );

        // Build debit and credit entries for Receipt and Contra
        XElement debitEntry_Receipt_Contra = new XElement("LEDGERENTRIES.LIST",
            new XElement("LEDGERNAME", ledger1),
            new XElement("ISDEEMEDPOSITIVE", "No"),
            new XElement("ISPARTYLEDGER", "Yes"),
            new XElement("AMOUNT", amount.ToString("F2"))
        );

        XElement creditEntry_Receipt_Contra = new XElement("LEDGERENTRIES.LIST",
            new XElement("LEDGERNAME", ledger2),
            new XElement("ISDEEMEDPOSITIVE", "No"),
            new XElement("ISPARTYLEDGER", "Yes"),
            new XElement("AMOUNT", (-amount).ToString("F2"))
        );


        // 🔄 Correct the order for display in Tally
        if (TransType.Equals("Receipt", StringComparison.OrdinalIgnoreCase))
        {
            // Debit (To) first, then Credit (By)
            voucherElement.Add(debitEntry_Receipt_Contra);
            voucherElement.Add(creditEntry_Receipt_Contra);
        }

        else if (TransType.Equals("Contra", StringComparison.OrdinalIgnoreCase))
        {
            // Debit (To) first, then Credit (By)
            voucherElement.Add(debitEntry_Receipt_Contra);
            voucherElement.Add(creditEntry_Receipt_Contra);
        }

        else if (TransType.Equals("Payment", StringComparison.OrdinalIgnoreCase))
        {
            // Credit (By) first, then Debit (To)
            voucherElement.Add(creditEntry_Payment_Journal);
            voucherElement.Add(debitEntry_Payment_Journal);
        }

        else if (TransType.Equals("Journal", StringComparison.OrdinalIgnoreCase))
        {
            // Credit (By) first, then Debit (To)
            voucherElement.Add(creditEntry_Payment_Journal);
            voucherElement.Add(debitEntry_Payment_Journal);
        }

        else
        {
            // Default order: Debit first
            voucherElement.Add(debitEntry_Payment_Journal);
            voucherElement.Add(creditEntry_Payment_Journal);
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
    private bool ValidateExcelData(DataTable table, int year1, int year2, out List<(int Row, string Date)> emptyDates, out List<(int Row, string Date)> invalidDateFormats, out List<(int Row, string Date)> incorrectYearDates, out List<int> noParticulars, out List<int> noValues, out List<int> noTransTypes)
    {
        string errorMessage = string.Empty;
        emptyDates = new List<(int Row, string Date)>();
        invalidDateFormats = new List<(int Row, string Date)>();
        incorrectYearDates = new List<(int Row, string Date)>();
        noParticulars = new List<int>();
        noValues = new List<int>();
        noTransTypes = new List<int>();

        if (!table.Columns.Contains("Date"))
        {
            errorMessage = "Excel file must contain a 'Date' column.";
            return false;
        }

        int rowIndex = 1; // Excel row index (1-based for user readability)
        foreach (DataRow row in table.Rows)
        {
            rowIndex++;

            // ===== Date Validation =====
            string dateStr = row["Date"]?.ToString();
            if (string.IsNullOrEmpty(dateStr))
            {
                emptyDates.Add((rowIndex, "Empty"));
            }
            else
            {
                DateTime date;
                bool parsed = false;

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
                }
                else
                {
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
                }
            }

            // ===== NoParticulars Check (Ledger 1 or Ledger 2 missing) =====
            string ledger1 = table.Columns.Contains("Ledger 1") ? row["Ledger 1"].ToString() : string.Empty;
            string ledger2 = table.Columns.Contains("Ledger 2") ? row["Ledger 2"].ToString() : string.Empty;
            if (string.IsNullOrWhiteSpace(ledger1) || string.IsNullOrWhiteSpace(ledger2))
            {
                noParticulars.Add(rowIndex);
            }

            // ===== NoValue Check (Amount missing or zero) =====
            string amountStr = table.Columns.Contains("Amount") ? row["Amount"].ToString() : string.Empty;
            if (!decimal.TryParse(amountStr, out decimal amount) || amount == 0)
            {
                noValues.Add(rowIndex);
            }

            // ===== NoVoucherType Check =====
            string voucherType = table.Columns.Contains("Transaction Type") ? row["Transaction Type"].ToString() : string.Empty;
            if (string.IsNullOrWhiteSpace(voucherType))
            {
                noTransTypes.Add(rowIndex);
            }
        }

        if (emptyDates.Any() || invalidDateFormats.Any() || incorrectYearDates.Any() || noParticulars.Any() || noValues.Any() || noTransTypes.Any())
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

                // Date Errors
                foreach (var error in emptyDates)
                    worksheet.Range(error.Row, 1, error.Row, lastColumn).Style.Fill.BackgroundColor = XLColor.Yellow;

                foreach (var error in invalidDateFormats)
                    worksheet.Range(error.Row, 1, error.Row, lastColumn).Style.Fill.BackgroundColor = XLColor.Blue;

                foreach (var error in incorrectYearDates)
                    worksheet.Range(error.Row, 1, error.Row, lastColumn).Style.Fill.BackgroundColor = XLColor.Red;

                // NoParticulars (Ledger 1 or Ledger 2 missing)
                foreach (var row in noParticulars)
                    worksheet.Range(row, 1, row, lastColumn).Style.Fill.BackgroundColor = XLColor.Orange;

                // NoValues (Amount missing or zero)
                foreach (var row in noValues)
                    worksheet.Range(row, 1, row, lastColumn).Style.Fill.BackgroundColor = XLColor.LightGreen;

                // NoVoucherType
                foreach (var row in noTransTypes)
                    worksheet.Range(row, 1, row, lastColumn).Style.Fill.BackgroundColor = XLColor.Pink;

                // Error Details Sheet
                var errorSheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == "Error Details")
                    ?? workbook.Worksheets.Add("Error Details");
                errorSheet.Clear();

                errorSheet.Cell(1, 1).Value = "Voucher Error Details";
                errorSheet.Cell(2, 1).Value = $"Generated On: {DateTime.Now:dd-MM-yyyy hh:mm:ss tt}";

                int logRow = 4;
                void LogList<T>(string title, IEnumerable<T> list)
                {
                    errorSheet.Cell(logRow, 1).Value = title;
                    int col = 2;
                    foreach (var item in list)
                        errorSheet.Cell(logRow, col++).Value = item.ToString();
                    logRow++;
                }

                LogList("Empty Dates (Yellow)", emptyDates.Select(e => $"Row {e.Row}"));
                LogList("Invalid Date Formats (Blue)", invalidDateFormats.Select(e => $"Row {e.Row}: {e.Date}"));
                LogList("Incorrect Year Dates (Red)", incorrectYearDates.Select(e => $"Row {e.Row}: {e.Date}"));
                LogList("No Ledger 1 or Ledger 2 (Orange)", noParticulars.Select(r => $"Row {r}"));
                LogList("Amount Missing or Zero (LightGreen)", noValues.Select(r => $"Row {r}"));
                LogList("Transaction Type Missing (Pink)", noTransTypes.Select(r => $"Row {r}"));

                // Auto-fit columns for better readability
                errorSheet.Columns().AdjustToContents();

                errorSheet.SetTabActive(false);

                if (worksheet.AutoFilter.IsEnabled)
                {
                    worksheet.AutoFilter.Clear();
                }
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
