using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Xml.Linq;
using ClosedXML.Excel;
using Microsoft.UI.Xaml;
using Office_Tools_Lite.Task_Helper;
using Windows.ApplicationModel.DataTransfer;
using Windows.Storage.Pickers;
using WinRT.Interop;

namespace Office_Tools_Lite.CLOSED_XML.XML_Generator;

public class Sales_Return_Excel_To_Xml_Converter
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
            string FinancialYear = await ShowDialog.ShowInputDialog("Sales Return XML Converter", mainWindow);
            if (string.IsNullOrEmpty(FinancialYear))
            {
                return; // User canceled, return safely
            }

            // Extract Financial Year and month
            string[] parts = FinancialYear.Contains("|") ? FinancialYear.Split('|') : new[] { "All", FinancialYear };
            string FYmonth = parts[0];
            string FY = parts.Length > 1 ? parts[1] : FinancialYear;
            string YN = parts[2].Trim();

            // Validate Financial Year format and extract years
            if (!ValidateFinancialYear(FY, out int year1, out int year2))
            {
                await ShowDialog.ShowMsgBox("Financial Year Error", $"Invalid Financial Year '{FY}'. Please use format YYYY-YY (e.g., 2024-25).", "OK", null, 1, mainWindow);
                return;
            }

            try
            {
                string excelPath = file.Path; // Use file.Path instead of dialog.FileName
                _table = await ReadExcel(excelPath);
                if (_table == null) return;

                // Prepare output lists for validation
                List<(int Row, string Date)> emptyDates;
                List<(int Row, string Date)> invalidDateFormats;
                List<(int Row, string Date)> incorrectYearDates;
                List<(int Row, string Date)> monthMismatchErrors;
                List<int> noBillNos;
                List<int> noNameBills;
                List<int> duplicateBillNos;
                List<int> noValueBills;
                List<int> negativeBills;
                List<int> grossMismatchBills;

                // Validate dates in the Excel data
                if (!ValidateExcelData(_table, year1, year2, FYmonth, out emptyDates, out invalidDateFormats, out incorrectYearDates, out monthMismatchErrors, out noBillNos, out noNameBills, out duplicateBillNos, out noValueBills, out negativeBills, out grossMismatchBills))
                {
                    await ShowDialog.ShowMsgBox("Date Validation Error", "There are few errors in Excel data...!", "OK", null, 1, mainWindow);
                    await HighlightErrorsInReadyDataSheetAsync(excelPath, emptyDates, invalidDateFormats, incorrectYearDates, monthMismatchErrors, noBillNos, noNameBills, noValueBills, negativeBills, duplicateBillNos, grossMismatchBills);
                    return;
                }

                XDocument xml = BuildTallyXml(_table, YN, FY);
                string outputPath = Path.Combine(Path.GetDirectoryName(excelPath), "Sales_Return_Vouchers.xml");
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
    private XDocument BuildTallyXml(DataTable table, string YN, string FY)
    {
        // 🔹 Create ledgers only once
        var ledgerMasters = table.AsEnumerable()
            .Select(r => r["Name"].ToString().Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Select(name =>
            {
                var row = table.AsEnumerable()
                                .First(r => r["Name"].ToString().Trim().Equals(name, StringComparison.OrdinalIgnoreCase));
                return createLedgerFromData(row, FY);
            });

        //var vouchers = table.AsEnumerable().Skip(1).Select(row => CreateVoucherElement(row, table));
        //var vouchers = table.AsEnumerable().Select(row => CreateVoucherElement(row, table));
        int maxRows = 15; // LIMIT to 15 rows
        var vouchers = table.AsEnumerable().Take(maxRows).Select(row => CreateVoucherElement(row, table));

        // 🔹 Decide REQUESTDATA based on Yes / No
        IEnumerable<XElement> requestDataElements = YN.Equals("Yes", StringComparison.OrdinalIgnoreCase)
                ? ledgerMasters.Concat(vouchers)   // YES → Masters + Vouchers
                : vouchers;

        var xml = new XDocument(
            new XElement("ENVELOPE",
                new XElement("HEADER",
                    new XElement("TALLYREQUEST", "Import Data")
                ),
                new XElement("BODY",
                    new XElement("IMPORTDATA",
                        new XElement("REQUESTDESC",
                            new XElement("REPORTNAME", "All Masters"),
                            new XElement("STATICVARIABLES",
                                new XElement("SVCURRENTCOMPANY", "")
                            )
                        ),
                        new XElement("REQUESTDATA", requestDataElements)
                    )
                )
            )
        );

        return xml;
    }
    #endregion

    #region Create Voucher Element
    private XElement CreateVoucherElement(DataRow row, DataTable table)
    {
        // Check for required columns
        string[] requiredColumns = { "Date", "Name", "Bill No.", "Total Invoice Value" };
        var missingColumns = requiredColumns.Where(col => !table.Columns.Contains(col)).ToList();
        if (missingColumns.Any())
        {
            throw new ArgumentException($"Missing required columns in Excel file: {string.Join(", ", missingColumns)}");
        }

        // Handle required columns
        string date = Convert.ToDateTime(row["Date"]).ToString("yyyyMMdd");
        string partyName = row["Name"].ToString();
        string voucherNumber = row["Bill No."].ToString();
        string grossTotal = decimal.Parse(row["Total Invoice Value"].ToString()).ToString("F2");

        // Handle optional column
        string referenceDate = "";
        if (table.Columns.Contains("Ref Date") && !string.IsNullOrEmpty(row["Ref Date"]?.ToString()))
        {
            referenceDate = Convert.ToDateTime(row["Ref Date"]).ToString("yyyyMMdd");
        }
        string referenceNumber = table.Columns.Contains("Ref No.") ? row["Ref No."].ToString() : ""; // Default to empty string if Ref No. is missing
        string narration = table.Columns.Contains("Narration") ? row["Narration"].ToString() : "";
        string type = table.Columns.Contains("Type") && !string.IsNullOrWhiteSpace(row["Type"]?.ToString()) ? $" ({row["Type"].ToString()})" : "";
        string vchType = table.Columns.Contains("Voucher Type") && !string.IsNullOrWhiteSpace(row["Voucher Type"]?.ToString()) ? row["Voucher Type"].ToString() : "Credit Note";

        //string gstin = table.Columns.Contains("GSTIN") ? row["GSTIN"].ToString() : "";

        XElement voucher = new XElement("TALLYMESSAGE",
            new XElement("VOUCHER",
                new XAttribute("VCHTYPE", vchType),
                new XAttribute("ACTION", "Create"),
                new XAttribute("OBJVIEW", "Invoice Voucher View"),
                new XElement("DATE", date),
                new XElement("VOUCHERTYPENAME", vchType),
                new XElement("VOUCHERNUMBER", voucherNumber),
                new XElement("PARTYNAME", partyName),
                new XElement("PARTYLEDGERNAME", partyName),
                new XElement("PLACEOFSUPPLY", "Karnataka"),
                new XElement("STATENAME", "Karnataka"),
                new XElement("COUNTRYOFRESIDENCE", "India"),
                new XElement("ISINVOICE", "Yes"),

                    new XElement("LEDGERENTRIES.LIST",
                    new XElement("LEDGERNAME", partyName),
                    new XElement("ISDEEMEDPOSITIVE", "No"),
                    new XElement("ISPARTYLEDGER", "Yes"),
                    new XElement("AMOUNT", grossTotal)
                )
            )
        );

        // Conditionally add PARTYGSTIN element if gstin is non-empty
        //if (!string.IsNullOrEmpty(gstin))
        //{
        //    voucher.Element("VOUCHER").Add(new XElement("PARTYGSTIN", gstin));
        //}
        if (!string.IsNullOrEmpty(narration))
        {
            voucher.Element("VOUCHER")?.Add(new XElement("NARRATION", narration));
        }

        if (!string.IsNullOrEmpty(referenceDate))
        {
            voucher.Element("VOUCHER")?.Add(new XElement("REFERENCEDATE", referenceDate));
        }

        if (!string.IsNullOrEmpty(referenceNumber))
        {
            voucher.Element("VOUCHER")?.Add(new XElement("REFERENCE", referenceNumber));
        }

        // Excel to Tally ledger name mapping
        var ledgerMapping = new Dictionary<string, string>
        {
            { "Non GST", "Sales Non GST Exempted" + $"{type}" },
            { "Sales A/c", "Sales A/c" + $"{type}" },
            { "Taxable 0%", "Sales GST @ 0%" + $"{type}" },

            { "Taxable 5%", "Sales GST @ 5%" + $"{type}" },
            { "Room Rent @ 5%", "Room Rent Collected @ 5%" + $"{type}" }, //Special type
            { "SGST 2.5%", "Output SGST @ 2.5%" + $"{type}" },
            { "CGST 2.5%", "Output CGST @ 2.5%" + $"{type}" },

            { "Taxable 12%", "Sales GST @ 12%" + $"{type}" },
            { "Room Rent @ 12%", "Room Rent Collected @ 12%" + $"{type}" }, //Special type
            { "SGST 6%", "Output SGST @ 6%" + $"{type}" },
            { "CGST 6%", "Output CGST @ 6%" + $"{type}" },

            { "Taxable 18%", "Sales GST @ 18%" + $"{type}" },
            { "Labour", "Service Charges @ 18%" + $"{type}" }, //Special type
            { "SGST 9%", "Output SGST @ 9%" + $"{type}" },
            { "CGST 9%", "Output CGST @ 9%" + $"{type}" },

            { "Taxable 28%", "Sales GST @ 28%" + $"{type}" },
            { "SGST 14%", "Output SGST @ 14%" + $"{type}" },
            { "CGST 14%", "Output CGST @ 14%" + $"{type}" },

            { "Inter State 0%", "Interstate Sales GST @ 0%" + $"{type}" },

            { "Inter State 5%", "Interstate Sales GST @ 5%" + $"{type}" },
            { "IGST 5%", "Output IGST @ 5%"+ $"{type}" },

            { "Inter State 12%", "Interstate Sales GST @ 12%" + $"{type}" },
            { "IGST 12%", "Output IGST @ 12%" + $"{type}" },

            { "Inter State 18%", "Interstate Sales GST @ 18%" + $"{type}" },
            { "IGST 18%", "Output IGST @ 18%" + $"{type}" },

            { "Inter State 28%", "Interstate Sales GST @ 28%" + $"{type}" },
            { "IGST 28%", "Output IGST @ 28%" + $"{type}" },

            { "CESS", "Output Cess @ 12%" + $"{type}" },
            { "Round Off", "Rounded Off" }
        };

        foreach (var kvp in ledgerMapping)
        {
            string colName = kvp.Key;
            string ledgerName = kvp.Value;

            if (table.Columns.Contains(colName))
            {
                string amount = row[colName].ToString();
                if (decimal.TryParse(amount, out decimal amt) && amt != 0)
                {
                    string finalAmount = amt > 0 ? (-amt).ToString("F2") : Math.Abs(amt).ToString("F2");

                    voucher.Element("VOUCHER").Add(
                        new XElement("LEDGERENTRIES.LIST",
                            new XElement("LEDGERNAME", ledgerName),
                            new XElement("ISDEEMEDPOSITIVE", "Yes"),
                            new XElement("ISPARTYLEDGER", "No"),
                            new XElement("AMOUNT", finalAmount)
                        )
                    );
                }
            }
        }

        return voucher;
    }
    #endregion

    #region Create ledger
    private static readonly Dictionary<string, string> GstStateMap = new()
    {
        { "01", "Jammu And Kashmir" },
        { "02", "Himachal Pradesh" },
        { "03", "Punjab" },
        { "04", "Chandigarh" },
        { "05", "Uttarakhand" },
        { "06", "Haryana" },
        { "07", "Delhi" },
        { "08", "Rajasthan" },
        { "09", "Uttar Pradesh" },
        { "10", "Bihar" },
        { "11", "Sikkim" },
        { "12", "Arunachal Pradesh" },
        { "13", "Nagaland" },
        { "14", "Manipur" },
        { "15", "Mizoram" },
        { "16", "Tripura" },
        { "17", "Meghlaya" },
        { "18", "Assam" },
        { "19", "West Bengal" },
        { "20", "Jharkhand" },
        { "21", "Odisha" },
        { "22", "Chattisgarh" },
        { "23", "Madhya Pradesh" },
        { "24", "Gujarat" },
        { "25", "Daman And Diu" },
        { "26", "Dadra And Nagar Haveli" },
        { "27", "Maharashtra" },
        { "28", "Andhra Pradesh(Before Division)" },
        { "29", "Karnataka" },
        { "30", "Goa" },
        { "31", "Lakshwadeep" },
        { "32", "Kerala" },
        { "33", "Tamil Nadu" },
        { "34", "Puducherry" },
        { "35", "Andaman And Nicobar Islands" },
        { "36", "Telangana" },
        { "37", "Andhra Pradesh" }
    };

    private string GetStateFromGstin(string gstin)
    {
        if (string.IsNullOrWhiteSpace(gstin) || gstin.Length < 2)
            return "";

        string code = gstin.Substring(0, 2);
        return GstStateMap.TryGetValue(code, out var state) ? state : "";
    }

    private XElement createLedgerFromData(DataRow row, string FY)
    {
        string name = row["Name"].ToString().Trim();
        string gstin = row.Table.Columns.Contains("GSTIN") ? row["GSTIN"].ToString().Trim() : "";
        string statecode = row.Table.Columns.Contains("My State Code") ? row["My State Code"].ToString().Trim() : "";
        string state = GetStateFromGstin(string.IsNullOrWhiteSpace(gstin) || gstin == "0" ? statecode : gstin);

        string country = "India";
        var parts = FY.Split('-');
        string applicablefrom = $"{parts[0]}0401";
        //string applicablefrom = "20170401";

        return new XElement("TALLYMESSAGE",
            new XElement("LEDGER",
                new XAttribute("NAME", name),
                new XAttribute("ACTION", "Create"),
                new XElement("NAME", name),
                new XElement("PARENT", "Sundry Creditors"),
                new XElement("GSTREGISTRATIONTYPE", !string.IsNullOrEmpty(gstin) && gstin != "0" ? "Regular" : "Unregistered"),
                new XElement("PARTYGSTIN", gstin),
                new XElement("ISGSTAPPLICABLE", !string.IsNullOrEmpty(gstin) && gstin != "0" ? "Yes" : "No"),
                new XElement("FBTCATEGORY", "Not Applicable"),
                new XElement("ISBILLWISEON", "No"),
                new XElement("ISCOSTCENTRESON", "No"),
                new XElement("ISINTERESTON", "No"),
                new XElement("USEFORGAINLOSS", "No"),
                new XElement("USEFORCOMPUTATION", "No"),
                new XElement("USEFORVAT", "No"),
                new XElement("IGNORETDSEXEMPT", "No"),
                new XElement("ISDEEMEDPOSITIVE", "No"),
                new XElement("AFFECTSSTOCK", "No"),
                new XElement("FORPAYROLL", "No"),
                new XElement("ISABCENABLED", "No"),
                new XElement("ISCOSTTRACKINGON", "No"),
                new XElement("STATENAME", state),

                //Applicable for Tally 6
                new XElement("LEDMAILINGDETAILS.LIST",
                new XElement("APPLICABLEFROM", applicablefrom),
                new XElement("MAILINGNAME", name),
                new XElement("STATE", state),
                new XElement("COUNTRY", country)),

                new XElement("LEDGSTREGDETAILS.LIST",
                new XElement("APPLICABLEFROM", applicablefrom),
                new XElement("GSTREGISTRATIONTYPE", !string.IsNullOrEmpty(gstin) && gstin != "0" ? "Regular" : "Unregistered"),
                new XElement("GSTIN", gstin))
            )
        );
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
    private bool ValidateExcelData(DataTable table, int year1, int year2, string FYmonth, out List<(int Row, string Date)> emptyDates, out List<(int Row, string Date)> invalidDateFormats, out List<(int Row, string Date)> incorrectYearDates, out List<(int Row, string Date)> monthMismatchErrors,
        out List<int> noBillNos, out List<int> noNameBills, out List<int> duplicateBillNos, out List<int> noValueBills, out List<int> negativeBills, out List<int> grossMismatchBills)
    {
        emptyDates = new List<(int Row, string Date)>();
        invalidDateFormats = new List<(int Row, string Date)>();
        incorrectYearDates = new List<(int Row, string Date)>();
        monthMismatchErrors = new List<(int Row, string Date)>();
        noBillNos = new List<int>();
        noNameBills = new List<int>();
        duplicateBillNos = new List<int>();
        noValueBills = new List<int>();
        negativeBills = new List<int>();
        grossMismatchBills = new List<int>();

        if (!table.Columns.Contains("Date"))
            return false;

        string[] months = { "All", "Apr", "May", "June", "July", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar" };
        int monthNumber = Array.IndexOf(months, FYmonth);

        Dictionary<int, int> monthMapping = new Dictionary<int, int>
            {
                { 4, 1 }, { 5, 2 }, { 6, 3 }, { 7, 4 }, { 8, 5 },
                { 9, 6 }, { 10, 7 }, { 11, 8 }, { 12, 9 },
                { 1, 10 }, { 2, 11 }, { 3, 12 }
            };

        HashSet<string> billNosTracker = new HashSet<string>();

        string[] taxableCols = {
                "Non GST", "Sales A/c", "Taxable 0%", "Taxable 5%", "Taxable 12%",
                "Taxable 18%", "Taxable 28%", "Inter State 0%", "Inter State 5%",
                "Inter State 12%", "Inter State 18%", "Inter State 28%",
                "Labour", "Room Rent @ 12%", "Room Rent @ 5%"
            };

        int rowIndex = 1;
        foreach (DataRow row in table.Rows)
        {
            rowIndex++;

            // Date Validation
            string dateStr = row["Date"] != null ? row["Date"].ToString() : string.Empty;
            if (string.IsNullOrEmpty(dateStr))
            {
                emptyDates.Add((rowIndex, "Empty"));
            }
            else
            {
                DateTime date;
                bool parsed = false;

                if (row["Date"] is DateTime)
                {
                    date = (DateTime)row["Date"];
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
                        dateStr = date.ToString("dd-MM-yyyy");
                }

                if (!parsed)
                {
                    invalidDateFormats.Add((rowIndex, dateStr));
                }
                else
                {
                    if (date.Month >= 4 && date.Month <= 12)
                    {
                        if (date.Year != year1)
                            incorrectYearDates.Add((rowIndex, dateStr));
                    }
                    else if (date.Month >= 1 && date.Month <= 3)
                    {
                        if (date.Year != year2)
                            incorrectYearDates.Add((rowIndex, dateStr));
                    }

                    if (FYmonth != "All" &&
                        monthMapping.ContainsKey(date.Month) &&
                        monthMapping[date.Month] != monthNumber)
                    {
                        monthMismatchErrors.Add((rowIndex, dateStr));
                    }
                }
            }

            // Bill No check
            string billNo = table.Columns.Contains("Bill No.") ? row["Bill No."].ToString() : string.Empty;
            decimal billNoDec;
            if (string.IsNullOrWhiteSpace(billNo) || (decimal.TryParse(billNo, out billNoDec) && billNoDec == 0))
                noBillNos.Add(rowIndex);

            // Name check
            string name = table.Columns.Contains("Name") ? row["Name"].ToString() : string.Empty;
            decimal nameDec;
            if (string.IsNullOrWhiteSpace(name) || (decimal.TryParse(name, out nameDec) && nameDec == 0) || name == "#NA" || name == "NoValueAvailable")
                noNameBills.Add(rowIndex);

            // Duplicate Bill No check
            if (!string.IsNullOrWhiteSpace(billNo))
            {
                if (!billNosTracker.Add(billNo))
                    duplicateBillNos.Add(rowIndex);
            }

            // Taxable value check
            decimal taxableTotal = 0;
            foreach (string col in taxableCols)
            {
                if (table.Columns.Contains(col))
                {
                    decimal val;
                    if (decimal.TryParse(row[col].ToString(), out val))
                        taxableTotal += val;
                }
            }
            if (taxableTotal == 0)
                noValueBills.Add(rowIndex);

            // Negative value check
            foreach (string col in taxableCols)
            {
                if (table.Columns.Contains(col))
                {
                    decimal val;
                    if (decimal.TryParse(row[col].ToString(), out val) && val < 0)
                    {
                        negativeBills.Add(rowIndex);
                        break;
                    }
                }
            }

            // --- Gross Total cross-check based on ledger mapping ---
            var ledgerKeysForGross = new List<string>
                {
                    "Non GST", "Sales A/c", "Taxable 0%", "Taxable 5%",
                    "SGST 2.5%", "CGST 2.5%",
                    "Taxable 12%", "SGST 6%", "CGST 6%",
                    "Taxable 18%", "SGST 9%", "CGST 9%",
                    "Taxable 28%", "SGST 14%", "CGST 14%",
                    "Inter State 0%", "Inter State 5%", "IGST 5%",
                    "Inter State 12%", "IGST 12%",
                    "Inter State 18%", "IGST 18%",
                    "Inter State 28%", "IGST 28%",
                    "CESS", "Labour", "Room Rent @ 12%", "Round Off"
                };

            decimal ledgerMappingTotal = 0;
            foreach (var key in ledgerKeysForGross)
            {
                if (table.Columns.Contains(key) &&
                    decimal.TryParse(row[key]?.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal val))
                {
                    ledgerMappingTotal += val;
                }
            }

            if (table.Columns.Contains("Total Invoice Value") &&
                decimal.TryParse(row["Total Invoice Value"]?.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal grossTotal))
            {
                if (Math.Round(grossTotal, 2) != Math.Round(ledgerMappingTotal, 2))
                {
                    grossMismatchBills.Add(rowIndex);
                }
            }


        }

        return !(emptyDates.Any() || invalidDateFormats.Any() || incorrectYearDates.Any() || monthMismatchErrors.Any() ||
                 noBillNos.Any() || noNameBills.Any() || duplicateBillNos.Any() ||
                 noValueBills.Any() || negativeBills.Any() || grossMismatchBills.Any());
    }

    #endregion

    #region Highlight Errors
    private async Task HighlightErrorsInReadyDataSheetAsync(string excelPath, List<(int Row, string Date)> emptyDates, List<(int Row, string Date)> invalidDateFormats,
                    List<(int Row, string Date)> incorrectYearDates, List<(int Row, string Date)> monthMismatchErrors, List<int> noBillNos, List<int> noNameBills, List<int> noValueBills,
                    List<int> negativeBills, List<int> duplicateBillNos, List<int> grossMismatchBills)
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

                // ----- Highlight rows in Ready Data -----
                // Yellow = Date-related issues
                foreach (var e in emptyDates)
                    worksheet.Range(e.Row, 1, e.Row, lastColumn).Style.Fill.BackgroundColor = XLColor.Yellow;
                foreach (var e in invalidDateFormats)
                    worksheet.Range(e.Row, 1, e.Row, lastColumn).Style.Fill.BackgroundColor = XLColor.Yellow;
                foreach (var e in incorrectYearDates)
                    worksheet.Range(e.Row, 1, e.Row, lastColumn).Style.Fill.BackgroundColor = XLColor.Yellow;
                foreach (var e in monthMismatchErrors)
                    worksheet.Range(e.Row, 1, e.Row, lastColumn).Style.Fill.BackgroundColor = XLColor.Yellow;

                // Orange = No taxable value
                foreach (var r in noValueBills)
                    worksheet.Range(r, 1, r, lastColumn).Style.Fill.BackgroundColor = XLColor.Orange;

                // Red = Missing bill no, missing name, negative values
                foreach (var r in noBillNos)
                    worksheet.Range(r, 1, r, lastColumn).Style.Fill.BackgroundColor = XLColor.Red;
                foreach (var r in noNameBills)
                    worksheet.Range(r, 1, r, lastColumn).Style.Fill.BackgroundColor = XLColor.Red;
                foreach (var r in negativeBills)
                    worksheet.Range(r, 1, r, lastColumn).Style.Fill.BackgroundColor = XLColor.Red;

                // LightBlue = Duplicate bill numbers
                foreach (var r in duplicateBillNos)
                    worksheet.Range(r, 1, r, lastColumn).Style.Fill.BackgroundColor = XLColor.LightBlue;

                // Pink = Gross total mismatch
                foreach (var r in grossMismatchBills)
                    worksheet.Range(r, 1, r, lastColumn).Style.Fill.BackgroundColor = XLColor.Pink;

                // ----- Error Details sheet (old layout style) -----
                var errorSheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == "Error Details") ??
                                 workbook.Worksheets.Add("Error Details");
                errorSheet.Clear();

                errorSheet.Cell(1, 1).Value = "Sales Voucher Error Details";
                errorSheet.Cell(2, 1).Value = "Generated On: " + DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss tt");

                int writeRow = 4;

                // helper to write tuple (Row,Date) lists horizontally: Title in col A, each error in subsequent columns
                Action<string, List<(int Row, string Date)>> writeTupleErrors = delegate (string title, List<(int Row, string Date)> list)
                {
                    errorSheet.Cell(writeRow, 1).Value = title;
                    int col = 2;
                    if (list == null || list.Count == 0)
                    {
                        // leave a placeholder so row is visible (optional)
                        errorSheet.Cell(writeRow, col).Value = "-";
                    }
                    else
                    {
                        foreach (var it in list)
                        {
                            errorSheet.Cell(writeRow, col++).Value = string.Format("Row {0}: {1}", it.Row, it.Date);
                        }
                    }
                    writeRow++;
                };

                // helper to write int lists horizontally: Title in col A, each row number in subsequent columns
                Action<string, List<int>> writeIntErrors = delegate (string title, List<int> list)
                {
                    errorSheet.Cell(writeRow, 1).Value = title;
                    int col = 2;
                    if (list == null || list.Count == 0)
                    {
                        errorSheet.Cell(writeRow, col).Value = "-";
                    }
                    else
                    {
                        foreach (var it in list)
                        {
                            errorSheet.Cell(writeRow, col++).Value = string.Format("Row {0}", it);
                        }
                    }
                    writeRow++;
                };

                // Write the various error categories (same order as old sheet)
                writeTupleErrors("Empty Dates", emptyDates);
                writeTupleErrors("Invalid Date Formats", invalidDateFormats);
                writeTupleErrors("Incorrect Year Dates", incorrectYearDates);
                writeTupleErrors("Month Mismatch Errors", monthMismatchErrors);

                writeIntErrors("No Bill Nos", noBillNos);
                writeIntErrors("No Name Bills", noNameBills);
                writeIntErrors("No Value Bills", noValueBills);
                writeIntErrors("Negative Value Bills", negativeBills);
                writeIntErrors("Duplicate Bill Nos", duplicateBillNos);
                writeIntErrors("Gross Mismatch Bills", grossMismatchBills);

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


}