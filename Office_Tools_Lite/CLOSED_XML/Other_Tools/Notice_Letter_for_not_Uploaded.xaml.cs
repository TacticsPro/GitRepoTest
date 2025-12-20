using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using ClosedXML.Excel;
using Windows.Storage;
using Windows.Storage.Pickers;
using Office_Tools_Lite.Task_Helper;

namespace Office_Tools_Lite.CLOSED_XML.Other_Tools;

public sealed partial class Notice_Letter_for_not_Uploaded : Microsoft.UI.Xaml.Controls.Page
{
    public Notice_Letter_for_not_Uploaded()
    {
        InitializeComponent();
    }

    private async void OnProceedClick(object sender, RoutedEventArgs e)
    {
        // Retrieve input values
        string firmName = FirmNameTextBox.Text;
        string firmGSTIN = FirmGstinTextBox.Text;
        string financialYear = FinancialYearTextBox.Text;
        string place = PlaceTextBox.Text;

        // Determine whether month-wise or yearly bills are selected
        string billsCombinedType = MonthWiseRadioButton.IsChecked == true ? "Month-Wise" : "Yearly";

        var requiredFields = new (TextBox, string)[]
        {
            (FirmNameTextBox, "Firm Name"),
            (FirmGstinTextBox, "Firm GSTIN"),
            (FinancialYearTextBox, "Financial Year"),
            (PlaceTextBox, "Place"),
        };

        foreach (var (field, name) in requiredFields)
        {
            if (string.IsNullOrWhiteSpace(field.Text))
            {
                await ShowDialog.ShowMsgBox("Warning", $" '{name}' cannot be empty.", "OK", null, 1, App.MainWindow);
                return;
            }
        }

        // Check if at least one radio button is selected  
        if (!MonthWiseRadioButton.IsChecked == true && !YearlyRadioButton.IsChecked == true)
        {
            await ShowDialog.ShowMsgBox("Warning", "Please select a 'Bills Combined Type'.", "OK", null, 1, App.MainWindow);
            return;
        }

        ProcessingText.Visibility = Visibility.Visible;

        // Prompt user to select Excel files
        FileOpenPicker picker = new FileOpenPicker();
        picker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
        picker.FileTypeFilter.Add(".xlsx");
        picker.FileTypeFilter.Add(".xls");

        // Ensure the picker works in desktop apps
        IntPtr hwnd = WinRT.Interop.WindowNative.GetWindowHandle(App.MainWindow);
        WinRT.Interop.InitializeWithWindow.Initialize(picker, hwnd);

        IReadOnlyList<StorageFile> files = await picker.PickMultipleFilesAsync();

        if (files.Count == 0)
        {
            await ShowDialog.ShowMsgBox("No Files Selected", "Please select at least one file.", "OK", null, 1, App.MainWindow);
            ProcessingText.Visibility = Visibility.Collapsed;
            return;
        }

        string outputFolderPath = string.Empty;

        foreach (var file in files)
        {
            // Retrieve the directory where the selected file is located
            StorageFolder fileDirectory = await file.GetParentAsync();

            // Create the Generated_Letters folder inside the file's directory
            StorageFolder generatedLettersFolder = await fileDirectory.CreateFolderAsync("Generated_Letters", CreationCollisionOption.OpenIfExists);

            // Generate the output file name
            string outputFileName = Path.GetFileNameWithoutExtension(file.Name) + "_letter.xlsx";
            StorageFile outputFile = await generatedLettersFolder.CreateFileAsync(outputFileName, CreationCollisionOption.ReplaceExisting);

            using (Stream inputStream = await file.OpenStreamForReadAsync())
            using (Stream outputStream = await outputFile.OpenStreamForWriteAsync())
            {
                // Generate the letter in Excel, passing the billsCombinedType to the method
                GenerateLetterInExcel(inputStream, outputStream, firmName, firmGSTIN, financialYear, place, billsCombinedType);

                // Ensure file stream is properly flushed
                outputStream.Flush();
            }

            // Log the output file path to debug
            string outputFilePath = outputFile.Path;

            // Store the directory of the output file
            outputFolderPath = Path.GetDirectoryName(outputFilePath);
        }

        await ShowDialog.ShowMsgBox("Success", $"Letter generated successfully! Files saved in: {outputFolderPath}", "OK", null, 1, App.MainWindow);

        ProcessingText.Visibility = Visibility.Collapsed;
        // Open the output folder (only once)
        if (!string.IsNullOrEmpty(outputFolderPath))
        {
            System.Diagnostics.Process.Start("explorer.exe", outputFolderPath);
        }
    }

    private void GenerateLetterInExcel(Stream inputStream, Stream outputStream, string firmName, string firmGSTIN, string financialYear, string place, string billsCombinedType)
    {
        // Load input workbook
        var inputWorkbook = new XLWorkbook(inputStream);
        var outputWorkbook = new XLWorkbook();

        foreach (var inputSheet in inputWorkbook.Worksheets)
        {
            // Skip empty sheets
            if (inputSheet.LastRowUsed() == null)
            {
                System.Diagnostics.Debug.WriteLine($"Skipping empty sheet: {inputSheet.Name}");
                continue;
            }

            string sheetName = inputSheet.Name;
            var outputSheet = outputWorkbook.Worksheets.Add($"Letter_{sheetName}");

            // Insert a new column at position "A" and set its width to 2
            outputSheet.Column(1).Width = 3;

            // Set column widths for multiple columns
            string[] columns = { "B", "C", "D", "E", "F", "G", "H", "I" };
            foreach (var col in columns)
            {
                outputSheet.Column(col).Width = 12.5;
            }

            // Retrieve "Name" and "GSTIN" from the input sheet
            string partyName = inputSheet.Cell("C2").GetString() ?? "Unknown";
            string partyGSTIN = inputSheet.Cell("D2").GetString() ?? "Unknown";

            // Merge cells and adjust font size as specified
            outputSheet.Range("B1:H1").Merge().Value = partyName;
            outputSheet.Range("B1:H1").Style.Font.Bold = true;
            outputSheet.Range("B1:H1").Style.Font.FontSize = 20;
            outputSheet.Range("B1:H1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            outputSheet.Range("B2:H2").Merge().Value = "Address -- here---";
            outputSheet.Range("B2:H2").Style.Font.FontSize = 11;
            outputSheet.Range("B2:H2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            outputSheet.Range("B3:H3").Merge().Value = "GST NO: " + partyGSTIN;
            outputSheet.Range("B3:H3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            outputSheet.Range("B5:H5").Merge().Value = "To Whom So Ever It May Concern";
            outputSheet.Range("B5:H5").Style.Font.Bold = true;
            outputSheet.Range("B5:H5").Style.Font.FontSize = 14;
            outputSheet.Range("B5:H5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            outputSheet.Range("B6:H10").Merge().Value = $"We {partyName}, bearing GSTIN- {partyGSTIN} hereby certify that we had " +
                $"supplied goods as mentioned in the invoices referenced below to {firmName} bearing GSTIN- {firmGSTIN}. Further, " +
                $"we wish to affirm that for the FY {financialYear}, the invoices mentioned below are declared in GSTR1 for the relevant " +
                $"Months/period in the B2C segment of GSTR1 instead of B2B Segment.";
            outputSheet.Range("B6:H10").Style.Font.FontSize = 12;
            outputSheet.Range("B6:H10").Style.Alignment.WrapText = true;
            outputSheet.Range("B6:H10").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            outputSheet.Row(6).Height = 30;

            // Additional logic to process invoices
            if (billsCombinedType == "Month-Wise")
            {
                AddMonthWiseInvoices(inputSheet, outputSheet, firmName, firmGSTIN, place);
            }
            else
            {
                AddYearlyInvoices(inputSheet, outputSheet, firmName, firmGSTIN, place);
            }
        }

        // Save the output workbook to the stream
        outputWorkbook.SaveAs(outputStream);
    }

    private void AddMonthWiseInvoices(IXLWorksheet dataSheet, IXLWorksheet outputSheet, string firmName, string firmGSTIN, string place)
    {
        // Headers for the invoice table
        var headers = new[] { "Invoice No", "Date", "Taxable Value", "CGST", "SGST", "Total" };
        var invoiceData = new List<object[]>();

        // Read data from input worksheet
        var lastRow = dataSheet.LastRowUsed()?.RowNumber() ?? 1;
        System.Diagnostics.Debug.WriteLine($"Last row used in {dataSheet.Name}: {lastRow}");

        for (int row = 2; row <= lastRow; row++)
        {
            var rowData = new object[8];
            for (int col = 1; col <= 8; col++)
            {
                rowData[col - 1] = dataSheet.Cell(row, col).Value;
            }
            invoiceData.Add(rowData);
        }

        if (!invoiceData.Any())
        {
            System.Diagnostics.Debug.WriteLine("No invoice data found in input sheet.");
            return;
        }

        string previousMonth = null;
        int rowIdx = 12; // Start row for invoice data

        foreach (var row in invoiceData)
        {
            // Handle invoice number
            var invoiceNo = row[0]?.ToString();
            if (string.IsNullOrEmpty(invoiceNo)) continue;

            // Handle date (try parsing if not DateTime)
            DateTime? date = null;
            if (row[1] is DateTime dt)
            {
                date = dt;
            }
            else if (row[1] != null && DateTime.TryParse(row[1].ToString(), out var parsedDate))
            {
                date = parsedDate;
            }
            if (date == null) continue;

            var monthName = date.Value.ToString("MMM");
            var taxableValue = row[4]?.ToString();
            var cgst = row[5]?.ToString();
            var sgst = row[6]?.ToString();
            var total = row[7]?.ToString();

            // Check if a new month needs a header
            if (monthName != previousMonth)
            {
                rowIdx += 1;
                outputSheet.Range($"B{rowIdx}:G{rowIdx}").Merge().Value = $"{monthName}";
                outputSheet.Range($"B{rowIdx}:G{rowIdx}").Style.Font.Bold = true;
                outputSheet.Range($"B{rowIdx}:G{rowIdx}").Style.Font.Underline = XLFontUnderlineValues.Single;
                outputSheet.Range($"B{rowIdx}:G{rowIdx}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                outputSheet.Range($"B{rowIdx}:G{rowIdx}").Style.Font.FontSize = 12;
                rowIdx += 1;

                // Headers for the invoice table
                for (int i = 0; i < headers.Length; i++)
                {
                    var cell = outputSheet.Cell(rowIdx, i + 2);
                    cell.Value = headers[i];
                    cell.Style.Font.Bold = true;
                    cell.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                }
                rowIdx += 1;
            }

            // Write invoice data
            if (double.TryParse(invoiceNo, out var numericInvoiceNo))
            {
                outputSheet.Cell(rowIdx, 2).Value = numericInvoiceNo;
            }
            else
            {
                outputSheet.Cell(rowIdx, 2).Value = invoiceNo;
            }

            outputSheet.Cell(rowIdx, 3).Value = date.Value.ToString("dd/MM/yyyy");
            outputSheet.Cell(rowIdx, 4).Value = Convert.ToDouble(taxableValue ?? "0");
            outputSheet.Cell(rowIdx, 5).Value = Convert.ToDouble(cgst ?? "0");
            outputSheet.Cell(rowIdx, 6).Value = Convert.ToDouble(sgst ?? "0");
            outputSheet.Cell(rowIdx, 7).Value = Convert.ToDouble(total ?? "0");

            // Apply border to each cell in the row
            for (int col = 2; col <= 7; col++)
            {
                outputSheet.Cell(rowIdx, col).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            }

            previousMonth = monthName;
            rowIdx += 1;
        }

        AddFooter(outputSheet, rowIdx + 1, firmName, place);
    }

    private void AddYearlyInvoices(IXLWorksheet dataSheet, IXLWorksheet outputSheet, string firmName, string firmGSTIN, string place)
    {
        // Headers for the invoice table
        var headers = new[] { "Invoice No", "Date", "Taxable Value", "CGST", "SGST", "Total" };
        var invoiceData = new List<object[]>();

        // Read data from input worksheet
        var lastRow = dataSheet.LastRowUsed()?.RowNumber() ?? 1;
        System.Diagnostics.Debug.WriteLine($"Last row used in {dataSheet.Name}: {lastRow}");

        for (int row = 2; row <= lastRow; row++)
        {
            var rowData = new object[8];
            for (int col = 1; col <= 8; col++)
            {
                rowData[col - 1] = dataSheet.Cell(row, col).Value;
            }
            invoiceData.Add(rowData);
        }

        if (!invoiceData.Any())
        {
            System.Diagnostics.Debug.WriteLine("No invoice data found in input sheet.");
            return;
        }

        // Add headers to the 12th row
        int headerRowIdx = 12;
        for (int i = 0; i < headers.Length; i++)
        {
            var cell = outputSheet.Cell(headerRowIdx, i + 2);
            cell.Value = headers[i];
            cell.Style.Font.Bold = true;
            cell.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
        }

        int rowIdx = 13;

        foreach (var row in invoiceData)
        {
            // Handle invoice number
            var invoiceNo = row[0]?.ToString();
            if (string.IsNullOrEmpty(invoiceNo)) continue;

            // Handle date (try parsing if not DateTime)
            DateTime? date = null;
            if (row[1] is DateTime dt)
            {
                date = dt;
            }
            else if (row[1] != null && DateTime.TryParse(row[1].ToString(), out var parsedDate))
            {
                date = parsedDate;
            }
            if (date == null) continue;

            var taxableValue = row[4]?.ToString();
            var cgst = row[5]?.ToString();
            var sgst = row[6]?.ToString();
            var total = row[7]?.ToString();

            // Write invoice data
            if (double.TryParse(invoiceNo, out var numericInvoiceNo))
            {
                outputSheet.Cell(rowIdx, 2).Value = numericInvoiceNo;
            }
            else
            {
                outputSheet.Cell(rowIdx, 2).Value = invoiceNo;
            }

            outputSheet.Cell(rowIdx, 3).Value = date.Value.ToString("dd/MM/yyyy");
            outputSheet.Cell(rowIdx, 4).Value = Convert.ToDouble(taxableValue ?? "0");
            outputSheet.Cell(rowIdx, 5).Value = Convert.ToDouble(cgst ?? "0");
            outputSheet.Cell(rowIdx, 6).Value = Convert.ToDouble(sgst ?? "0");
            outputSheet.Cell(rowIdx, 7).Value = Convert.ToDouble(total ?? "0");

            // Apply border to each cell in the row
            for (int i = 0; i < headers.Length; i++)
            {
                outputSheet.Cell(rowIdx, i + 2).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            }
            rowIdx += 1;
        }

        AddFooter(outputSheet, rowIdx + 1, firmName, place);
    }


    private void AddFooter(IXLWorksheet letterSheet, int lastRow, string firmName, string place)
    {
        // Set the footer text with appropriate line breaks
        letterSheet.Range($"B{lastRow + 1}:H{lastRow + 3}").Merge().Value =
            $"This certificate is issued at the request of {firmName}." +
            "to enable eligibility for availed input tax credit due to mismatch in GSTR2A under section 16 (2) " +
            "Read with Rule 36 & 37 of Karnataka GST Act and Central GST Act pursuant to Circular No. 183/15/2022.";
        letterSheet.Range($"B{lastRow + 1}:H{lastRow + 3}").Style.Alignment.WrapText = true;

        // Add the date and place
        letterSheet.Cell(lastRow + 5, 2).Value = $"DATE: {DateTime.Now:dd-MM-yyyy}";
        letterSheet.Cell(lastRow + 5, 7).Value = $"Yours Faithfully";
        letterSheet.Cell(lastRow + 6, 2).Value = $"Place: {place}";
    }

}