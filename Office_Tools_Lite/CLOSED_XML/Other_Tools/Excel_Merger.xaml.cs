using ClosedXML.Excel;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Office_Tools_Lite.Task_Helper;
using Windows.Storage.Pickers;

namespace Office_Tools_Lite.CLOSED_XML.Other_Tools;

public sealed partial class Excel_Merger : Page
{
    public Excel_Merger()
    {
        InitializeComponent();
    }

    private async void ProceedButton_Click(object sender, RoutedEventArgs e)
    {
        int optionSheets = GetSelectedOption();

        if (optionSheets == 0)
        {
            await ShowDialog.ShowMsgBox("Warning", "Select Required field", "OK", null, 1, App.MainWindow);
            return;
        }

        string sheetName = string.Empty;

        if (optionSheets == 2)
        {
            var inputBox = new TextBox
            {
                PlaceholderText = "Enter sheet name",
                Width = 300
            };

            var dialog = await ShowDialog.ShowMsgBox( "Selected Sheet", inputBox, "OK", "Cancel", 1, App.MainWindow);

            if (dialog == ContentDialogResult.Primary)
            {
                sheetName = inputBox.Text;
            }
            else
            {
                ProcessingText.Visibility = Visibility.Collapsed;
                return;
            }
        }

        var picker = new FileOpenPicker();
        picker.FileTypeFilter.Add(".xlsx"); // ClosedXML supports only .xlsx
        picker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;

        var hwnd = WinRT.Interop.WindowNative.GetWindowHandle(App.MainWindow);
        WinRT.Interop.InitializeWithWindow.Initialize(picker, hwnd);

        ProcessingText.Visibility = Visibility.Visible;

        var files = await picker.PickMultipleFilesAsync();

        if (files == null || !files.Any())
        {
            ProcessingText.Visibility = Visibility.Collapsed;
            return;
        }

        string[] excelFiles = files.Select(f => f.Path).ToArray();

        if (optionSheets == 1)
        {
            await MergeAllSheets(excelFiles);
        }
        else
        {
            await MergeSelectedSheet(excelFiles, sheetName);
        }

        ProcessingText.Visibility = Visibility.Collapsed;
    }

    private int GetSelectedOption()
    {
        if (AllSheets.IsChecked == true)
            return 1;
        if (SelectedSheet.IsChecked == true)
            return 2;
        if (ActiveSheet.IsChecked == true)
            return 3;
        return 0;
    }

    public async Task MergeAllSheets(string[] excelFiles)
    {
        string originalDir = Path.GetDirectoryName(excelFiles[0]);
        string subDir = Path.Combine(originalDir, "Output_Merged_Files");
        Directory.CreateDirectory(subDir);

        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("MergedSheet");
            int currentRow = 1;

            foreach (var excelFile in excelFiles)
            {
                using (var sourceWorkbook = new XLWorkbook(excelFile))
                {
                    foreach (var sheet in sourceWorkbook.Worksheets)
                    {
                        var usedRange = sheet.RangeUsed();
                        if (usedRange != null)
                        {
                            // Copy the used range to the target worksheet
                            var rowCount = usedRange.RowCount();
                            var colCount = usedRange.ColumnCount();

                            for (int row = 1; row <= rowCount; row++)
                            {
                                for (int col = 1; col <= colCount; col++)
                                {
                                    var cellValue = usedRange.Cell(row, col).Value;
                                    worksheet.Cell(currentRow + row - 1, col).Value = cellValue;
                                    // Preserve basic formatting (optional)
                                    worksheet.Cell(currentRow + row - 1, col).Style = usedRange.Cell(row, col).Style;
                                }
                            }
                            currentRow += rowCount;
                        }
                    }
                }
            }

            string outputExcelFile = Path.Combine(subDir, "Merged_File_All_Sheets.xlsx");
            workbook.SaveAs(outputExcelFile);

            await ShowDialog.ShowMsgBox("Merge All Sheets", "Successfully Merged Selected Sheets.", "OK", null, 1, App.MainWindow);

            var outputFolderPath = Path.GetDirectoryName(outputExcelFile);
            System.Diagnostics.Process.Start("explorer.exe", outputFolderPath);
        }
    }

    private async Task MergeSelectedSheet(string[] excelFiles, string sheetNamePatterns)
    {
        string originalDir = Path.GetDirectoryName(excelFiles[0]);
        string subDir = Path.Combine(originalDir, "Output_Files");
        Directory.CreateDirectory(subDir);

        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("MergedSheet");
            int currentRow = 1;

            // Split the input patterns (e.g., "B2B,CDNR" -> ["B2B", "CDNR"])
            var patterns = sheetNamePatterns.Split(',', (char)StringSplitOptions.RemoveEmptyEntries)
                                           .Select(p => p.Trim().ToLower())
                                           .ToArray();

            foreach (var excelFile in excelFiles)
            {
                using (var sourceWorkbook = new XLWorkbook(excelFile))
                {
                    // Find sheets matching any of the patterns
                    var sheetsToCopy = sourceWorkbook.Worksheets
                        .Where(ws => patterns.Any(p => ws.Name.ToLower().Contains(p)))
                        .ToList();

                    if (!sheetsToCopy.Any() && patterns.Length == 0)
                    {
                        sheetsToCopy.Add(sourceWorkbook.Worksheets.First()); // Fallback to first sheet if no patterns provided
                    }

                    foreach (var sheetToCopy in sheetsToCopy)
                    {
                        var usedRange = sheetToCopy.RangeUsed();
                        if (usedRange != null)
                        {
                            var rowCount = usedRange.RowCount();
                            var colCount = usedRange.ColumnCount();

                            for (int row = 1; row <= rowCount; row++)
                            {
                                for (int col = 1; col <= colCount; col++)
                                {
                                    var cellValue = usedRange.Cell(row, col).Value;
                                    worksheet.Cell(currentRow + row - 1, col).Value = cellValue;
                                    worksheet.Cell(currentRow + row - 1, col).Style = usedRange.Cell(row, col).Style;
                                }
                            }
                            currentRow += rowCount;
                        }
                    }
                }
            }

            string outputExcelFile = Path.Combine(subDir, "Merged_File_Selected_Sheets.xlsx");
            workbook.SaveAs(outputExcelFile);

            await ShowDialog.ShowMsgBox("Merge Selected Sheet", "Successfully Merged Selected Sheets from Selected Files.", "OK", null, 1, App.MainWindow);

            var outputFolderPath = Path.GetDirectoryName(outputExcelFile);
            System.Diagnostics.Process.Start("explorer.exe", outputFolderPath);
        }
    }
}