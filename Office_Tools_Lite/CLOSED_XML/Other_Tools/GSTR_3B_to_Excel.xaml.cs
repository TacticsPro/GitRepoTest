using System.Diagnostics.CodeAnalysis;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Office_Tools_Lite.Task_Helper;

namespace Office_Tools_Lite.CLOSED_XML.Other_Tools;

public sealed partial class GSTR_3B_to_Excel : Page
{
    public GSTR_3B_to_Excel()
    {
        this.InitializeComponent();
        
    }

    //[DynamicDependency(DynamicallyAccessedMemberTypes.All, typeof(ComboBoxItem))]
    //[DynamicDependency(DynamicallyAccessedMemberTypes.All, typeof(PdfReader))]
    //[DynamicDependency(DynamicallyAccessedMemberTypes.All, typeof(PdfDocument))]
    //private static void PreserveIText()
    //{
    //}
    private string GetComboBoxTag(ComboBox combo, string fallback)
    {
        if (combo.SelectedItem is ComboBoxItem item && item.Tag is string tag)
            return tag;
        return fallback;
    }

    #region Proceed
    private async void ProceedBtn_Click(object sender, RoutedEventArgs e)
    {
        ProceedBtn.IsEnabled = false;
        ProcessingText.Visibility = Visibility.Visible;

        //var tax_type = (TaxTypeCombo.SelectedItem as ComboBoxItem)?.Content.ToString();
        //var each_month_excel_files = (YesNoCombo.SelectedItem as ComboBoxItem)?.Content.ToString();

        string tax_type = GetComboBoxTag(TaxTypeCombo, "Total");
        string each_month_excel_files = GetComboBoxTag(YesNoCombo, "No");


        // Open File Picker to select multiple PDF files
        var files = await Filepicks();
        if (files == null || files.Count == 0)
        {
            ProceedBtn.IsEnabled = true;
            ProcessingText.Visibility = Visibility.Collapsed;
            return;
        }

        var ExtractedData = new List<Dictionary<string, object>>();

       
        foreach (var file in files)
        {
            var extractedData = await ExtractGSTR3BDetails(file);
            ExtractedData.Add(extractedData);

            // Export individual PDFs to separate Excel files
            if (each_month_excel_files == "Yes")
            {
                await ExportToExcel(extractedData, file.Name);
            }

        }

        if (each_month_excel_files == "Yes")
        {
            await ShowDialog.ShowMsgBox("Export Complete", $"Each month Excel file exported successfully", "Ok", null, 1, App.MainWindow);
        }

        ///Now export consolidated Excel after all PDFs are processed
        if (tax_type == "Total")
        {
            await ConsolidateExport_Total(ExtractedData);
            ProceedBtn.IsEnabled = true;
            ProcessingText.Visibility = Visibility.Collapsed;
        }
        else
        {
           await ConsolidateExport_Seperate(ExtractedData);
            ProceedBtn.IsEnabled = true;
            ProcessingText.Visibility = Visibility.Collapsed;
        }

        
    }
    #endregion

    #region File Pick
    private async Task<List<FileInfo>> Filepicks()
    {
        var hwnd = WinRT.Interop.WindowNative.GetWindowHandle(App.MainWindow);
        var picker = new Windows.Storage.Pickers.FileOpenPicker();

        // Initialize the picker with the window handle (important for WinUI 3)
        WinRT.Interop.InitializeWithWindow.Initialize(picker, hwnd);

        picker.SuggestedStartLocation = Windows.Storage.Pickers.PickerLocationId.DocumentsLibrary;
        picker.FileTypeFilter.Add(".pdf");

        var files = await picker.PickMultipleFilesAsync();

        if (files == null || files.Count == 0)
        {
            await ShowDialog.ShowMsgBox("Warning", "No PDF files were selected.", "Ok", null, 1, App.MainWindow);
            ProceedBtn.IsEnabled = true;
            return null;
        }

        // Convert StorageFiles to FileInfo objects
        return files.Select(file => new FileInfo(file.Path)).ToList();
    }
    #endregion

    #region Extract GSTR_3B Details
    private async Task<Dictionary<string, object>> ExtractGSTR3BDetails(FileInfo file)
    {
        // Defensive check to avoid native crash
        if (!File.Exists(file.FullName) || new FileInfo(file.FullName).Length < 100)
        {
            await ShowDialog.ShowMsgBox("Error", $"PDF file '{file.Name}' is missing or too small to parse.", "Ok", null, 1, App.MainWindow);
            return new Dictionary<string, object>();
        }

        var extractedData = new Dictionary<string, object>
            {
                {"Basic Details", new Dictionary<string, string>()},
                {"Supply Details", new Dictionary<string, Dictionary<string, double>>()},
                {"ITC Details", new Dictionary<string, Dictionary<string, double>>()},
                {"Exempt & Non-GST Supplies", new Dictionary<string, Dictionary<string, double>>()}
            };

        string pdfText = await Task.Run(() =>
        {
            try
            {
                using var fs = new FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var pdfReader = new PdfReader(fs);
                using var pdfDocument = new PdfDocument(pdfReader);

                var text = "";
                for (int i = 1; i <= pdfDocument.GetNumberOfPages(); i++)
                {
                    text += PdfTextExtractor.GetTextFromPage(pdfDocument.GetPage(i));
                }
                return text;
            }
            catch (Exception ex)
            {
                return $"PDF parsing error: {ex.Message}";
            }
        });



        #region Regex patterns for Basic Details
        // Regex patterns for Basic Details
        var basicDetailsPatterns = new Dictionary<string, string>
            {
                {"Year", @"Year\s+(\d{4}-\d{2})"},
                {"Period", @"Period\s+(\w+)"},
                {"GSTIN", @"GSTIN of the supplier\s+([A-Z0-9@]+)"},
                {"Legal name", @"2\(a\)\. Legal name of the registered person\s+([^\n]+)"},
                {"Trade name", @"2\(b\)\. Trade name, if any\s+([^\n]+)"},
                {"Date of ARN", @"2\(d\)\. Date of ARN\s+(\d{2}/\d{2}/\d{4})"}
            };

        foreach (var pattern in basicDetailsPatterns)
        {
            var match = Regex.Match(pdfText, pattern.Value);
            ((Dictionary<string, string>)extractedData["Basic Details"])[pattern.Key] = match.Success ? match.Groups[1].Value : "Not found";
        }

        // Extract the year and month from Basic Details
        int calendarYear = 0;
        int monthIndex = 0;
        var yearMatch = Regex.Match(pdfText, basicDetailsPatterns["Year"]);
        if (yearMatch.Success)
        {
            var yearParts = yearMatch.Groups[1].Value.Split('-');
            if (yearParts.Length == 2 && int.TryParse(yearParts[0], out var startYear))
            {
                // Parse the second year (e.g., '25' -> 2025)
                if (int.TryParse(yearParts[1], out var endYearPart))
                {
                    int endYear = 2000 + endYearPart; // e.g., '25' -> 2025
                    var monthMatch = Regex.Match(pdfText, basicDetailsPatterns["Period"]);
                    if (monthMatch.Success)
                    {
                        var month = monthMatch.Groups[1].Value;
                        var monthMap = new Dictionary<string, int>
                        {
                            {"January", 1}, {"February", 2}, {"March", 3}, {"April", 4},
                            {"May", 5}, {"June", 6}, {"July", 7}, {"August", 8},
                            {"September", 9}, {"October", 10}, {"November", 11}, {"December", 12}
                        };
                        monthIndex = monthMap.ContainsKey(month) ? monthMap[month] : 0;

                        // Determine calendar year: January–March use endYear, April–December use startYear
                        calendarYear = (monthIndex >= 1 && monthIndex <= 3) ? endYear : startYear;
                    }
                }
            }
        }

        #endregion Regex patterns for Basic Details

        #region Regex patterns for Supply Details
        // Regex patterns for Supply Details
        var supplyDetailsPatterns = new Dictionary<string, string>
            {
                {"Outward taxable supplies", @"\(a\) Outward taxable supplies \(other than zero rated, nil rated and[^\d]*(\d+(?:,\d+)*\.\d+)\s+(\d+(?:,\d+)*\.\d+)\s+(\d+(?:,\d+)*\.\d+)\s+(\d+(?:,\d+)*\.\d+)\s+(\d+(?:,\d+)*\.\d+)"},
                {"Outward taxable supplies (zero rated)", @"\(b\) Outward taxable supplies \(zero rated\)[^\d]*(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?"},
                {"Other outward supplies", @"\(c\) Other outward supplies \(nil rated, exempted\)[^\d]*(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?"},
                {"Inward supplies", @"\(d\) Inward supplies \(liable to reverse charge\)[^\d]*(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?"},
                {"Non-GST outward supplies", @"\(e\) Non-GST outward supplies[^\d]*(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?"}
            };



        foreach (var pattern in supplyDetailsPatterns)
        {
            var match = Regex.Match(pdfText, pattern.Value, RegexOptions.Singleline);
            if (match.Success)
            {
                ((Dictionary<string, Dictionary<string, double>>)extractedData["Supply Details"])[pattern.Key] = new Dictionary<string, double>
                    {
                        {"Total Taxable Value", double.TryParse(match.Groups[1]?.Value.Replace(",", ""), out var value1) ? value1 : 0.0},
                        {"Integrated Tax", double.TryParse(match.Groups[2]?.Value.Replace(",", ""), out var value2) ? value2 : 0.0},
                        {"Central Tax", double.TryParse(match.Groups[3]?.Value.Replace(",", ""), out var value3) ? value3 : 0.0},
                        {"State/UT Tax", double.TryParse(match.Groups[4]?.Value.Replace(",", ""), out var value4) ? value4 : 0.0},
                        {"Cess", double.TryParse(match.Groups[5]?.Value.Replace(",", ""), out var value5) ? value5 : 0.0}
                    };
            }
            else
            {
                ((Dictionary<string, Dictionary<string, double>>)extractedData["Supply Details"])[pattern.Key] = new Dictionary<string, double>
                    {
                        {"Total Taxable Value", 0.0},
                        {"Integrated Tax", 0.0},
                        {"Central Tax", 0.0},
                        {"State/UT Tax", 0.0},
                        {"Cess", 0.0}
                    };
            }
        }

        #endregion Regex patterns for Supply Details

        #region Regex patterns for ITC Details
        // Regex patterns for ITC Details
        var itcDetailsPatterns = new Dictionary<string, string>
            {
                {"Inward supplies liable to reverse charge", @"\(3\) Inward supplies liable to reverse charge \(other than 1 & 2 above\)[^\d]*(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?"},
                {"All other ITC", @"\(5\) All other ITC[^\d]*(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?"},
                {"ITC Reversed - As per rules", @"\(1\) As per rules 38,42 & 43 of CGST Rules and section 17\(5\)[^\d]*(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?"},
                {"ITC Reversed - Others", @"\(2\) Others[^\d]*(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?"},
                {"Net ITC available", @"C\. Net ITC available \(A-B\)[^\d]*(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?"}
            };

        foreach (var pattern in itcDetailsPatterns)
        {
            var match = Regex.Match(pdfText, pattern.Value, RegexOptions.Singleline);
            if (match.Success)
            {
                ((Dictionary<string, Dictionary<string, double>>)extractedData["ITC Details"])[pattern.Key] = new Dictionary<string, double>
                    {
                        {"Integrated Tax", double.TryParse(match.Groups[1]?.Value.Replace(",", ""), out var value1) ? value1 : 0.0},
                        {"Central Tax", double.TryParse(match.Groups[2]?.Value.Replace(",", ""), out var value2) ? value2 : 0.0},
                        {"State/UT Tax", double.TryParse(match.Groups[3]?.Value.Replace(",", ""), out var value3) ? value3 : 0.0},
                        {"Cess", double.TryParse(match.Groups[4]?.Value.Replace(",", ""), out var value4) ? value4 : 0.0}
                    };
            }
            else
            {
                ((Dictionary<string, Dictionary<string, double>>)extractedData["ITC Details"])[pattern.Key] = new Dictionary<string, double>
                    {
                        {"Integrated Tax", 0.0},
                        {"Central Tax", 0.0},
                        {"State/UT Tax", 0.0},
                        {"Cess", 0.0}
                    };
            }
        }

        #endregion Regex patterns for ITC Details

        #region Regex patterns for Exempt and Non-GST Supplies
        // Regex patterns for Exempt and Non-GST Supplies
        var exemptPatterns = new Dictionary<string, string>
            {
                {"From a supplier under composition scheme, Exempt, Nil rated supply", @"From a supplier under composition scheme, Exempt, Nil rated supply\s+(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?"},
                {"Non GST supply", @"Non GST supply\s+(\d+(?:,\d+)*\.\d+)?\s+(\d+(?:,\d+)*\.\d+)?"}
            };

        foreach (var pattern in exemptPatterns)
        {
            var match = Regex.Match(pdfText, pattern.Value, RegexOptions.Singleline);
            if (match.Success)
            {
                ((Dictionary<string, Dictionary<string, double>>)extractedData["Exempt & Non-GST Supplies"])[pattern.Key] = new Dictionary<string, double>
                    {
                        {"Inter-State Supplies", double.TryParse(match.Groups[1]?.Value.Replace(",", ""), out var value1) ? value1 : 0.0},
                        {"Intra-State Supplies", double.TryParse(match.Groups[2]?.Value.Replace(",", ""), out var value2) ? value2 : 0.0}
                    };
            }
            else
            {
                ((Dictionary<string, Dictionary<string, double>>)extractedData["Exempt & Non-GST Supplies"])[pattern.Key] = new Dictionary<string, double>
                    {
                        {"Inter-State Supplies", 0.0},
                        {"Intra-State Supplies", 0.0}
                    };
            }
        }
        #endregion Regex patterns for Exempt and Non-GST Supplies

        #region Regex pattern for extracting Section (A) Other than reverse charge

        // Regex pattern for extracting Section (A) Other than reverse charge

        var paymentPatternA = @"\(A\) Other than reverse charge\s*\n(.*?)\n\(B\) Reverse charge";
        var matchA = Regex.Match(pdfText, paymentPatternA, RegexOptions.Singleline);

        if (matchA.Success)
        {
            var paymentSectionA = matchA.Groups[1].Value;

            if (calendarYear < 2024 || (calendarYear == 2024 && monthIndex <= 8))
            {

                #region 9 Columns for Other than Reverse Charge

                // Updated regex to match tax payment details with separate ITC columns
                var taxPattern = @"(Integrated|Central tax|State/UT tax|Cess)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)";
                var paymentMatches = Regex.Matches(paymentSectionA, taxPattern);

                extractedData["Payment Details"] = new Dictionary<string, Dictionary<string, double>>();

                foreach (Match match in paymentMatches)
                {
                    var taxType = match.Groups[1].Value;
                    if (taxType == "Integrated")
                        taxType = "Integrated tax";
                    else if (taxType == "Central")
                        taxType = "Central tax";
                    else if (taxType == "State/UT")
                        taxType = "State/UT tax";
                    else if (taxType == "Cess")
                        taxType = "Cess";

                    // Sum the three ITC columns
                    double taxPaidThroughITC = (double.TryParse(match.Groups[3].Value, out var itcIntegrated) ? itcIntegrated : 0.0) +
                                               (double.TryParse(match.Groups[4].Value, out var itcCentral) ? itcCentral : 0.0) +
                                               (double.TryParse(match.Groups[5].Value, out var itcStateUT) ? itcStateUT : 0.0) +
                                               (double.TryParse(match.Groups[6].Value, out var itcCess) ? itcCess : 0.0);

                    ((Dictionary<string, Dictionary<string, double>>)extractedData["Payment Details"])[taxType] = new Dictionary<string, double>
                        {
                            {"Tax paid through ITC", taxPaidThroughITC},
                            {"Tax paid in cash", double.TryParse(match.Groups[7].Value, out var taxPaidCash) ? taxPaidCash : 0.0},
                            {"Interest paid in cash", double.TryParse(match.Groups[8].Value, out var interestPaidCash) ? interestPaidCash : 0.0},
                            {"Late fee paid in cash", double.TryParse(match.Groups[9].Value, out var lateFeePaidCash) ? lateFeePaidCash : 0.0}
                        };

                }
                #endregion 9 Columns for Other than Reverse Charge
            }
            else
            {
            #region 11 Columns for Other than Reverse Charge
            //// Updated regex to match tax payment details with separate ITC columns
            var taxPattern = @"(Integrated|Central|State/UT|Cess)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)";
            var paymentMatches = Regex.Matches(paymentSectionA, taxPattern);

            extractedData["Payment Details"] = new Dictionary<string, Dictionary<string, double>>();

            foreach (Match match in paymentMatches)
            {
                var taxType = match.Groups[1].Value;
                if (taxType == "Integrated")
                    taxType = "Integrated tax";
                else if (taxType == "Central")
                    taxType = "Central tax";
                else if (taxType == "State/UT")
                    taxType = "State/UT tax";
                else if (taxType == "Cess")
                    taxType = "Cess";

                // Sum the three ITC columns
                double taxPaidThroughITC = (double.TryParse(match.Groups[5].Value, out var itcIntegrated) ? itcIntegrated : 0.0) +
                                            (double.TryParse(match.Groups[6].Value, out var itcCentral) ? itcCentral : 0.0) +
                                            (double.TryParse(match.Groups[7].Value, out var itcStateUT) ? itcStateUT : 0.0) +
                                            (double.TryParse(match.Groups[8].Value, out var itcCess) ? itcCess : 0.0);

                ((Dictionary<string, Dictionary<string, double>>)extractedData["Payment Details"])[taxType] = new Dictionary<string, double>
                    {
                        {"Tax paid through ITC", taxPaidThroughITC},
                        {"Tax paid in cash", double.TryParse(match.Groups[9].Value, out var taxPaidCash) ? taxPaidCash : 0.0},
                        {"Interest paid in cash", double.TryParse(match.Groups[10].Value, out var interestPaidCash) ? interestPaidCash : 0.0},
                        {"Late fee paid in cash", double.TryParse(match.Groups[11].Value, out var lateFeePaidCash) ? lateFeePaidCash : 0.0}
                    };

            }

            #endregion 11 Columns for Other than Reverse Charge

            }

        }

        #endregion Regex pattern for extracting Section (A) Other than reverse charge

        #region Regex pattern for extracting Section (B) Reverse charge
        // Regex pattern for extracting Section (B) Reverse charge
        var paymentPatternB = @"\(B\) Reverse charge(?: and supplies made u/s 9\(5\))?\s*\n(.*?)\nBreakup of tax liability declared";
        var matchB = Regex.Match(pdfText, paymentPatternB, RegexOptions.Singleline);

        if (matchB.Success)
        {
            var paymentSectionB = matchB.Groups[1].Value;

            if (calendarYear < 2024 || (calendarYear == 2024 && monthIndex <= 8))
            {
                #region 9 Columns for Reverse Charge
                // Corrected regex to extract tax payment details
                var taxPattern = @"(Integrated|Central tax|State/UT tax|Cess)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)";
                var paymentMatches = Regex.Matches(paymentSectionB, taxPattern);

                extractedData["Reverse Charge Payment Details"] = new Dictionary<string, Dictionary<string, double>>();

                foreach (Match match in paymentMatches)
                {
                    var taxType = match.Groups[1].Value;
                    // Replace the switch expression with a switch-case
                    switch (taxType)
                    {
                        case "Integrated":
                            taxType = "Integrated tax";
                            break;
                        case "Central":
                            taxType = "Central tax";
                            break;
                        case "State/UT":
                            taxType = "State/UT tax";
                            break;
                        case "Cess":
                            taxType = "Cess";
                            break;
                        default:
                            break;
                    }

                    ((Dictionary<string, Dictionary<string, double>>)extractedData["Reverse Charge Payment Details"])[taxType] = new Dictionary<string, double>
                    {
                            {"Tax paid through ITC", double.TryParse(match.Groups[5].Value, out var value1) ? value1 : 0.0}, // not used because no ITC for reverse charge
                            {"Tax paid in cash", double.TryParse(match.Groups[7].Value, out var value2) ? value2 : 0.0},
                            {"Interest paid in cash", double.TryParse(match.Groups[8].Value, out var value3) ? value3 : 0.0},
                            {"Late fee paid in cash", double.TryParse(match.Groups[9].Value, out var value4) ? value4 : 0.0}
                    };
                }
                #endregion 9 columns
            }
            else
            {
                #region 11 Columns for Reverse Charge

                // Corrected regex to extract tax payment details
                var taxPattern = @"(Integrated|Central|State/UT|Cess)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.-]+)";
                var paymentMatches = Regex.Matches(paymentSectionB, taxPattern);

                extractedData["Reverse Charge Payment Details"] = new Dictionary<string, Dictionary<string, double>>();

                foreach (Match match in paymentMatches)
                {
                    var taxType = match.Groups[1].Value;
                    // Replace the switch expression with a switch-case
                    switch (taxType)
                    {
                        case "Integrated":
                            taxType = "Integrated tax";
                            break;
                        case "Central":
                            taxType = "Central tax";
                            break;
                        case "State/UT":
                            taxType = "State/UT tax";
                            break;
                        case "Cess":
                            taxType = "Cess";
                            break;
                        default:
                            break;
                    }

                    ((Dictionary<string, Dictionary<string, double>>)extractedData["Reverse Charge Payment Details"])[taxType] = new Dictionary<string, double>
                    {
                            {"Tax paid through ITC", double.TryParse(match.Groups[5].Value, out var value1) ? value1 : 0.0},  //not used because no ITC for reverse charge
                            {"Tax paid in cash", double.TryParse(match.Groups[9].Value, out var value2) ? value2 : 0.0},
                            {"Interest paid in cash", double.TryParse(match.Groups[10].Value, out var value3) ? value3 : 0.0},
                            {"Late fee paid in cash", double.TryParse(match.Groups[11].Value, out var value4) ? value4 : 0.0}
                    };
                }
                #endregion 11 Columns for Reverse charge
            }

        }

        #endregion Regex pattern for extracting Section (B) Reverse charge

        return extractedData;
    }
    #endregion

    #region Export to Excel
    private async Task ExportToExcel(Dictionary<string, object> extractedData, string pdfFileName)
    {
        string excelFileName = Path.GetFileNameWithoutExtension(pdfFileName) + "_GSTR-3B.xlsx";
        string excelFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), excelFileName);

        using (var workbook = new XLWorkbook())
        {
            // Add a single worksheet
            var worksheet = workbook.Worksheets.Add("GSTR-3B Details");

            int row = 1;

            // Write Basic Details
            if (extractedData.ContainsKey("Basic Details"))
            {
                worksheet.Cell(row, 1).Value = "Basic Details";
                worksheet.Cell(row, 1).Style.Font.Bold = true;
                row++;

                var basicDetails = (Dictionary<string, string>)extractedData["Basic Details"];
                foreach (var item in basicDetails)
                {
                    worksheet.Cell(row, 1).Value = item.Key;
                    worksheet.Cell(row, 2).Value = item.Value;
                    row++;
                }
                row++; // Add an empty row for spacing
            }

            // Write Supply Details
            if (extractedData.ContainsKey("Supply Details"))
            {
                worksheet.Cell(row, 1).Value = "Supply Details";
                worksheet.Range(row, 1, row, 6).Merge(); // Merge cells for header
                worksheet.Cell(row, 1).Style.Font.Bold = true;
                row++;

                // Add column headers
                worksheet.Cell(row, 1).Value = "Description";
                worksheet.Cell(row, 2).Value = "Total Taxable Value";
                worksheet.Cell(row, 3).Value = "Integrated Tax";
                worksheet.Cell(row, 4).Value = "Central Tax";
                worksheet.Cell(row, 5).Value = "State/UT Tax";
                worksheet.Cell(row, 6).Value = "Cess";
                worksheet.Range(row, 1, row,6).Style.Font.Bold = true;


                row++;

                var supplyDetails = (Dictionary<string, Dictionary<string, double>>)extractedData["Supply Details"];
                foreach (var item in supplyDetails)
                {
                    worksheet.Cell(row, 1).Value = item.Key;
                    worksheet.Cell(row, 2).Value = item.Value["Total Taxable Value"];
                    worksheet.Cell(row, 3).Value = item.Value["Integrated Tax"];
                    worksheet.Cell(row, 4).Value = item.Value["Central Tax"];
                    worksheet.Cell(row, 5).Value = item.Value["State/UT Tax"];
                    worksheet.Cell(row, 6).Value = item.Value["Cess"];
                    row++;
                }
                row++; // Add an empty row for spacing
            }

            // Write ITC Details
            if (extractedData.ContainsKey("ITC Details"))
            {
                worksheet.Cell(row, 1).Value = "ITC Details";
                worksheet.Range(row, 1, row, 5).Merge(); // Merge cells for header
                worksheet.Cell(row, 1).Style.Font.Bold = true;
                row++;

                // Add column headers
                worksheet.Cell(row, 1).Value = "Description";
                worksheet.Cell(row, 2).Value = "Integrated Tax";
                worksheet.Cell(row, 3).Value = "Central Tax";
                worksheet.Cell(row, 4).Value = "State/UT Tax";
                worksheet.Cell(row, 5).Value = "Cess Tax";
                worksheet.Range(row, 1, row, 5).Style.Font.Bold = true;

                row++;

                var itcDetails = (Dictionary<string, Dictionary<string, double>>)extractedData["ITC Details"];
                foreach (var item in itcDetails)
                {
                    worksheet.Cell(row, 1).Value = item.Key;
                    worksheet.Cell(row, 2).Value = item.Value["Integrated Tax"];
                    worksheet.Cell(row, 3).Value = item.Value["Central Tax"];
                    worksheet.Cell(row, 4).Value = item.Value["State/UT Tax"];
                    worksheet.Cell(row, 5).Value = item.Value["Cess"];
                    row++;
                }
                row++; // Add an empty row for spacing
            }

            // Write Exempt & Non-GST Supplies
            if (extractedData.ContainsKey("Exempt & Non-GST Supplies"))
            {
                worksheet.Cell(row, 1).Value = "Exempt & Non-GST Supplies";
                worksheet.Range(row, 1, row, 3).Merge(); // Merge cells for header
                worksheet.Cell(row, 1).Style.Font.Bold = true;
                row++;

                // Add column headers
                worksheet.Cell(row, 1).Value = "Description";
                worksheet.Cell(row, 2).Value = "Inter-State";
                worksheet.Cell(row, 3).Value = "Intra-State";
                worksheet.Range(row, 1, row, 3).Style.Font.Bold = true;

                row++;

                var exemptSupplies = (Dictionary<string, Dictionary<string, double>>)extractedData["Exempt & Non-GST Supplies"];
                foreach (var item in exemptSupplies)
                {
                    worksheet.Cell(row, 1).Value = item.Key;
                    worksheet.Cell(row, 2).Value = item.Value["Inter-State Supplies"];
                    worksheet.Cell(row, 3).Value = item.Value["Intra-State Supplies"];
                    row++;
                }
                row++; // Add an empty row for spacing
            }

            // Write Tax Payment Details
            if (extractedData.ContainsKey("Payment Details"))
            {
                worksheet.Cell(row, 1).Value = "Tax Payment Details (A)";
                worksheet.Range(row, 1, row, 5).Merge(); // Merge cells for header
                worksheet.Cell(row, 1).Style.Font.Bold = true;
                row++;

                // Add column headers
                worksheet.Cell(row, 1).Value = "Description";
                worksheet.Cell(row, 2).Value = "Tax paid through ITC";
                worksheet.Cell(row, 3).Value = "Tax paid in cash";
                worksheet.Cell(row, 4).Value = "Interest paid in cash";
                worksheet.Cell(row, 5).Value = "Late fee paid in cash";
                worksheet.Range(row, 1, row, 5).Style.Font.Bold = true;

                row++;

                var paymentDetails = (Dictionary<string, Dictionary<string, double>>)extractedData["Payment Details"];
                foreach (var item in paymentDetails)
                {
                    worksheet.Cell(row, 1).Value = item.Key;
                    worksheet.Cell(row, 2).Value = item.Value["Tax paid through ITC"];
                    worksheet.Cell(row, 3).Value = item.Value["Tax paid in cash"];
                    worksheet.Cell(row, 4).Value = item.Value["Interest paid in cash"];
                    worksheet.Cell(row, 5).Value = item.Value["Late fee paid in cash"];
                    row++;
                }
                row++; // Add an empty row for spacing
            }

            // Write Tax Payment Details (B)
            if (extractedData.ContainsKey("Reverse Charge Payment Details"))
            {
                worksheet.Cell(row, 1).Value = "Tax Payment Details (B)";
                worksheet.Range(row, 1, row, 5).Merge(); // Merge cells for header
                worksheet.Cell(row, 1).Style.Font.Bold = true;
                row++;

                // Add column headers for (B)
                worksheet.Cell(row, 1).Value = "Description";
                worksheet.Cell(row, 2).Value = "Tax paid through ITC";
                worksheet.Cell(row, 3).Value = "Tax paid in cash";
                worksheet.Cell(row, 4).Value = "Interest paid in cash";
                worksheet.Cell(row, 5).Value = "Late fee paid in cash";
                worksheet.Range(row, 1, row, 5).Style.Font.Bold = true;

                row++;

                var reverseChargeDetails = (Dictionary<string, Dictionary<string, double>>)extractedData["Reverse Charge Payment Details"];
                foreach (var item in reverseChargeDetails)
                {
                    worksheet.Cell(row, 1).Value = item.Key;
                    worksheet.Cell(row, 2).Value = item.Value.ContainsKey("Tax paid through ITC") ? item.Value["Tax paid through ITC"] : 0.0;
                    worksheet.Cell(row, 3).Value = item.Value.ContainsKey("Tax paid in cash") ? item.Value["Tax paid in cash"] : 0.0;
                    worksheet.Cell(row, 4).Value = item.Value.ContainsKey("Interest paid in cash") ? item.Value["Interest paid in cash"] : 0.0;
                    worksheet.Cell(row, 5).Value = item.Value.ContainsKey("Late fee paid in cash") ? item.Value["Late fee paid in cash"] : 0.0;
                    row++;
                }
                row++; // Add an empty row for spacing
            }
            // Determine the last row and last column dynamically
            int lastRow = worksheet.LastRowUsed().RowNumber();
            int lastColumn = worksheet.LastColumnUsed().ColumnNumber();

            // Apply borders to only the exported data
            var dataRange = worksheet.Range(1, 1, lastRow, lastColumn); // Select the range with data
            dataRange.Style.Border.TopBorder = XLBorderStyleValues.Thin;
            dataRange.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            dataRange.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            dataRange.Style.Border.RightBorder = XLBorderStyleValues.Thin;

            // Auto-fit columns for better readability
            worksheet.Columns().AdjustToContents();

            workbook.SaveAs(excelFilePath);

            //await ShowDialog.ShowMsgBox1("Export Complete", $"Excel file exported successfully to: {excelFilePath}");

        }
    }

    #endregion

    #region Consolidated Export_Totla
    private async Task ConsolidateExport_Total(List<Dictionary<string, object>> extractedDataList)
    {
        string excelFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Combined_GSTR3B_1.xlsx");

        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Summary");
            int row = 1;

            // Add headers
            worksheet.Cell(row, 1).Value = "Period";
            worksheet.Cell(row, 2).Value = "Date of ARN";
            worksheet.Cell(row, 3).Value = "Name";
            worksheet.Cell(row, 4).Value = "Total Taxable Value";
            worksheet.Cell(row, 5).Value = "Total Tax (Outward)";
            worksheet.Cell(row, 6).Value = "Net ITC";
            worksheet.Cell(row, 7).Value = "Total Exempted Purchase";
            worksheet.Cell(row, 8).Value = "Tax paid in cash (A)";
            worksheet.Cell(row, 9).Value = "Interest paid in cash (A)";
            worksheet.Cell(row, 10).Value = "Late fee paid in cash (A)";
            worksheet.Cell(row, 11).Value = "Tax paid in cash (B)";
            worksheet.Cell(row, 12).Value = "Interest paid in cash (B)";
            worksheet.Cell(row, 13).Value = "Late fee paid in cash (B)";
            worksheet.Range(row, 1, row, 13).Style.Font.Bold = true;
            row++;

            foreach (var extractedData in extractedDataList) // Loop through all PDFs
            {
                var basicDetails = (Dictionary<string, string>)extractedData["Basic Details"];
                string period = basicDetails.ContainsKey("Period") ? basicDetails["Period"] : "Unknown";
                string arnDate = basicDetails.ContainsKey("Date of ARN") ? basicDetails["Date of ARN"] : "Unknown";
                string tradename = basicDetails.ContainsKey("Trade name") ? basicDetails["Trade name"] : "Unknown";

                worksheet.Cell(row, 1).Value = period;
                worksheet.Cell(row, 2).Value = arnDate;
                worksheet.Cell(row, 3).Value = tradename;

                var supplyDetails = (Dictionary<string, Dictionary<string, double>>)extractedData["Supply Details"];
                if (supplyDetails.ContainsKey("Outward taxable supplies"))
                {
                    var outward = supplyDetails["Outward taxable supplies"];
                    double totalOutwardTax = outward["Integrated Tax"] + outward["Central Tax"] + outward["State/UT Tax"] + outward["Cess"];
                    worksheet.Cell(row, 4).Value = outward["Total Taxable Value"];
                    worksheet.Cell(row, 5).Value = totalOutwardTax;
                }

                var itcDetails = (Dictionary<string, Dictionary<string, double>>)extractedData["ITC Details"];
                if (itcDetails.ContainsKey("Net ITC available"))
                {
                    var itc = itcDetails["Net ITC available"];
                    double totalITC = itc["Integrated Tax"] + itc["Central Tax"] + itc["State/UT Tax"] + itc["Cess"];
                    worksheet.Cell(row, 6).Value = totalITC;
                }
                //
                double total_exempted = 0;
            
                var exemptDetails = (Dictionary<string, Dictionary<string, double>>)extractedData["Exempt & Non-GST Supplies"];
                foreach (var details in exemptDetails.Values)
                {
                    total_exempted += details["Intra-State Supplies"];
                    total_exempted += details["Inter-State Supplies"];

                }
                worksheet.Cell(row, 7).Value = total_exempted;
                //
                double totalTaxPaidA = 0, interestPaidA = 0, lateFeePaidA = 0;
                double totalTaxPaidB = 0, interestPaidB = 0, lateFeePaidB = 0;

                var paymentDetails = (Dictionary<string, Dictionary<string, double>>)extractedData["Payment Details"];
                foreach (var details in paymentDetails.Values)
                {
                    totalTaxPaidA += details["Tax paid in cash"];
                    interestPaidA += details["Interest paid in cash"];
                    lateFeePaidA += details["Late fee paid in cash"];
                }

                var reverseChargeDetails = (Dictionary<string, Dictionary<string, double>>)extractedData["Reverse Charge Payment Details"];
                foreach (var details in reverseChargeDetails.Values)
                {
                    totalTaxPaidB += details["Tax paid in cash"];
                    interestPaidB += details["Interest paid in cash"];
                    lateFeePaidB += details["Late fee paid in cash"];
                }

                worksheet.Cell(row, 8).Value = totalTaxPaidA;
                worksheet.Cell(row, 9).Value = interestPaidA;
                worksheet.Cell(row, 10).Value = lateFeePaidA;
                worksheet.Cell(row, 11).Value = totalTaxPaidB;
                worksheet.Cell(row, 12).Value = interestPaidB;
                worksheet.Cell(row, 13).Value = lateFeePaidB;

                row++; // Move to next row for the next file
            }
            // Determine the last row and last column dynamically
            int lastRow = worksheet.LastRowUsed().RowNumber();
            int lastColumn = worksheet.LastColumnUsed().ColumnNumber();

            // Apply borders to only the exported data
            var dataRange = worksheet.Range(1, 1, lastRow, lastColumn); // Select the range with data
            dataRange.Style.Border.TopBorder = XLBorderStyleValues.Thin;
            dataRange.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            dataRange.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            dataRange.Style.Border.RightBorder = XLBorderStyleValues.Thin;


            // Auto-fit columns for better readability
            worksheet.Columns().AdjustToContents();

            
            workbook.SaveAs(excelFilePath);

            await ShowDialog.ShowMsgBox("Export Complete", $"Consolidated Excel file exported successfully to: {excelFilePath}","Ok", null, 1, App.MainWindow);

            /// Open the output folder (only once)
            string outputfolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            if (!string.IsNullOrEmpty(outputfolder))
            {
                System.Diagnostics.Process.Start("explorer.exe", outputfolder);
            }
        }
    }

    #endregion

    #region Consolidated Export_Seperate
    private async Task ConsolidateExport_Seperate(List<Dictionary<string, object>> extractedDataList)
    {
        string showTax = TaxTypeCombo.Text.ToString();
        string excelFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Combined_GSTR3B_2.xlsx");

        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Summary");

            int row = 1;

            worksheet.Cell(row, 4).Value = "Outward";
            worksheet.Cell(row, 9).Value = "Net ITC";
            worksheet.Cell(row, 14).Value = "Other than Reverse charge (A)";
            worksheet.Cell(row, 25).Value = "Reverse charge (B)";
            worksheet.Range(row, 1, row, 37).Style.Font.Bold = true;
            worksheet.Range(row+1, 1, row+1, 37).Style.Font.Bold = true;
            
            // Merge cells D1 to H1
            worksheet.Range(1, 4, 1, 8).Merge();
            worksheet.Cell(1, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; // Center align text
            worksheet.Cell(1, 4).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center; // Center align vertically
            // Merge cells I1 to L1
            worksheet.Range(1, 9, 1, 12).Merge();
            worksheet.Cell(1, 9).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; // Center align text
            worksheet.Cell(1, 9).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center; // Center align vertically
            // Merge cells N1 to Y1
            worksheet.Range(1, 14, 1, 25).Merge();
            worksheet.Cell(1, 14).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left; // Left align text
            worksheet.Cell(1, 14).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center; // Center align vertically
            // Merge cells Z1 to AK1
            worksheet.Range(1, 26, 1, 37).Merge();
            worksheet.Cell(1, 26).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left; // Left align text
            worksheet.Cell(1, 26).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center; // Center align vertically

            row++;

            worksheet.Cell(row, 1).Value = "Period";
            worksheet.Cell(row, 2).Value = "Date of ARN";
            worksheet.Cell(row, 3).Value = "Name";
            worksheet.Cell(row, 4).Value = "Total Taxable Value";
            worksheet.Cell(row, 5).Value = "Integrated Tax (Outward)";
            worksheet.Cell(row, 6).Value = "Central Tax (Outward)";
            worksheet.Cell(row, 7).Value = "State/UT Tax (Outward)";
            worksheet.Cell(row, 8).Value = "Cess (Outward)";

            worksheet.Cell(row, 9).Value = "Integrated Tax (ITC)";
            worksheet.Cell(row, 10).Value = "Central Tax (ITC)";
            worksheet.Cell(row, 11).Value = "State/UT Tax (ITC)";
            worksheet.Cell(row, 12).Value = "Cess (ITC)";
            worksheet.Cell(row, 13).Value = "Total Exempted Purchase";
            worksheet.Cell(row, 14).Value = "Tax paid in cash (Integrated) A";
            worksheet.Cell(row, 15).Value = "Tax paid in cash (Central) A";
            worksheet.Cell(row, 16).Value = "Tax paid in cash (State/UT) A";
            worksheet.Cell(row, 17).Value = "Tax paid in cash (Cess) A";
            worksheet.Cell(row, 18).Value = "Interest paid in cash (Integrated) A";
            worksheet.Cell(row, 19).Value = "Interest paid in cash (Central) A";
            worksheet.Cell(row, 20).Value = "Interest paid in cash (State/UT) A";
            worksheet.Cell(row, 21).Value = "Interest paid in cash (Cess) A";
            worksheet.Cell(row, 22).Value = "Late fee paid in cash (Integrated) A";
            worksheet.Cell(row, 23).Value = "Late fee paid in cash (Central) A";
            worksheet.Cell(row, 24).Value = "Late fee paid in cash (State/UT) A";
            worksheet.Cell(row, 25).Value = "Late fee paid in cash (Cess) A";

            worksheet.Cell(row, 26).Value = "Tax paid in cash (Integrated) B";
            worksheet.Cell(row, 27).Value = "Tax paid in cash (Central) B";
            worksheet.Cell(row, 28).Value = "Tax paid in cash (State/UT) B";
            worksheet.Cell(row, 29).Value = "Tax paid in cash (Cess) B";
            worksheet.Cell(row, 30).Value = "Interest paid in cash (Integrated) B";
            worksheet.Cell(row, 31).Value = "Interest paid in cash (Central) B";
            worksheet.Cell(row, 32).Value = "Interest paid in cash (State/UT) B";
            worksheet.Cell(row, 33).Value = "Interest paid in cash (Cess) B";
            worksheet.Cell(row, 34).Value = "Late fee paid in cash (Integrated) B";
            worksheet.Cell(row, 35).Value = "Late fee paid in cash (Central) B";
            worksheet.Cell(row, 36).Value = "Late fee paid in cash (State/UT) B";
            worksheet.Cell(row, 37).Value = "Late fee paid in cash (Cess) B";
            row++;

            foreach (var extractedData in extractedDataList) // Loop through all PDFs
            {
                var basicDetails = (Dictionary<string, string>)extractedData["Basic Details"];
                string period = basicDetails.ContainsKey("Period") ? basicDetails["Period"] : "Unknown";
                string arnDate = basicDetails.ContainsKey("Date of ARN") ? basicDetails["Date of ARN"] : "Unknown";
                string tradename = basicDetails.ContainsKey("Trade name") ? basicDetails["Trade name"] : "Unknown";

                worksheet.Cell(row, 1).Value = period;
                worksheet.Cell(row, 2).Value = arnDate;
                worksheet.Cell(row, 3).Value = tradename;

                var supplyDetails = (Dictionary<string, Dictionary<string, double>>)extractedData["Supply Details"];
                if (supplyDetails.ContainsKey("Outward taxable supplies"))
                {
                    var outward = supplyDetails["Outward taxable supplies"];
                    worksheet.Cell(row, 4).Value = outward["Total Taxable Value"];
                    worksheet.Cell(row, 5).Value = outward["Integrated Tax"];
                    worksheet.Cell(row, 6).Value = outward["Central Tax"];
                    worksheet.Cell(row, 7).Value = outward["State/UT Tax"];
                    worksheet.Cell(row, 8).Value = outward["Cess"];
                }

                var itcDetails = (Dictionary<string, Dictionary<string, double>>)extractedData["ITC Details"];
                if (itcDetails.ContainsKey("Net ITC available"))
                {
                    var itc = itcDetails["Net ITC available"];
                    worksheet.Cell(row, 9).Value = itc["Integrated Tax"];
                    worksheet.Cell(row, 10).Value = itc["Central Tax"];
                    worksheet.Cell(row, 11).Value = itc["State/UT Tax"];
                    worksheet.Cell(row, 12).Value = itc["Cess"];
                }

                double total_exempted = 0;

                var exemptDetails = (Dictionary<string, Dictionary<string, double>>)extractedData["Exempt & Non-GST Supplies"];
                foreach (var details in exemptDetails.Values)
                {
                    total_exempted += details["Intra-State Supplies"];
                    total_exempted += details["Inter-State Supplies"];


                }
                worksheet.Cell(row, 13).Value = total_exempted;
                //
                double totalTaxPaidA = 0, interestPaidA = 0, lateFeePaidA = 0;
                double totalTaxPaidB = 0, interestPaidB = 0, lateFeePaidB = 0;

                var paymentDetails = (Dictionary<string, Dictionary<string, double>>)extractedData["Payment Details"];
                foreach (var taxType in paymentDetails.Keys)
                {
                    var details = paymentDetails[taxType];
                    worksheet.Cell(row, 14 + (taxType == "Integrated tax" ? 0 : taxType == "Central tax" ? 1 : taxType == "State/UT tax" ? 2 : 3)).Value = details["Tax paid in cash"];
                    worksheet.Cell(row, 18 + (taxType == "Integrated tax" ? 0 : taxType == "Central tax" ? 1 : taxType == "State/UT tax" ? 2 : 3)).Value = details["Interest paid in cash"];
                    worksheet.Cell(row, 22 + (taxType == "Integrated tax" ? 0 : taxType == "Central tax" ? 1 : taxType == "State/UT tax" ? 2 : 3)).Value = details["Late fee paid in cash"];
                }

                var reverseChargeDetails = (Dictionary<string, Dictionary<string, double>>)extractedData["Reverse Charge Payment Details"];
                foreach (var taxType in reverseChargeDetails.Keys)
                {
                    var details = reverseChargeDetails[taxType];
                    worksheet.Cell(row, 26 + (taxType == "Integrated tax" ? 0 : taxType == "Central tax" ? 1 : taxType == "State/UT tax" ? 2 : 3)).Value = details["Tax paid in cash"];
                    worksheet.Cell(row, 30 + (taxType == "Integrated tax" ? 0 : taxType == "Central tax" ? 1 : taxType == "State/UT tax" ? 2 : 3)).Value = details["Interest paid in cash"];
                    worksheet.Cell(row, 34 + (taxType == "Integrated tax" ? 0 : taxType == "Central tax" ? 1 : taxType == "State/UT tax" ? 2 : 3)).Value = details["Interest paid in cash"];
                }
                row++;

                // Determine the last row and last column dynamically
                int lastRow = worksheet.LastRowUsed().RowNumber();
                int lastColumn = worksheet.LastColumnUsed().ColumnNumber();

                // Apply borders to only the exported data
                var dataRange = worksheet.Range(1, 1, lastRow, lastColumn); // Select the range with data
                dataRange.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                dataRange.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                dataRange.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                dataRange.Style.Border.RightBorder = XLBorderStyleValues.Thin;


                // Auto-fit columns for better readability
                worksheet.Columns().AdjustToContents();
                
                workbook.SaveAs(excelFilePath);
            }
            await ShowDialog.ShowMsgBox("Export Complete",$"Consolidated Excel file exported successfully to: {excelFilePath}", "Ok", null, 1, App.MainWindow);

            // Open the output folder (only once)
            string outputfolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            if (!string.IsNullOrEmpty(outputfolder))
            {
                System.Diagnostics.Process.Start("explorer.exe", outputfolder);
            }
        }
    }
    #endregion

}

