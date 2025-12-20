using System.Runtime.InteropServices;
using ClosedXML.Excel;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Office_Tools_Lite.Task_Helper;
using Windows.Storage;
using WinRT.Interop;

namespace Office_Tools_Lite.CLOSED_XML.Data_Entry;

public class Voucher_No_Correction
{
    public KeyboardSimulator keyboard;
    public const int TimeInterval = 300; // 100 milliseconds = 0.1 Seconds // 1000 miliseconds = 1 second
    public const int Relax = 15000; // 100 milliseconds = 0.1 Seconds // 1000 miliseconds = 1 second

    public Voucher_No_Correction()
    {
        keyboard = new KeyboardSimulator();
    }

    #region Execute
    public async Task Execute()
    {
        var dialog = await ShowDialog.ShowMsgBox(
            "Voucher No. Correction",
            "Press Ok to continue.",
            "OK", "Cancel", 1, App.MainWindow);

        if (dialog != ContentDialogResult.Primary)
        {
            return;
        }

        // Open File Picker to select a single Excel file
        var file = await Filepick();
        if (file == null)
        {
            return;
        }

        // Read data from the Excel file
        var exceldata = await ReadData(file);
        if (exceldata == null || !exceldata.Any())
        {
            return;
        }

        //******** using seperate  Timer Window ********//
        TimerWindow timerWindow = new();
        timerWindow.Show();
        Window_Handler.Minimize(App.MainWindow);
        await timerWindow.StartTimer(5);

        // Attempt to activate the TallyPrime window
        bool isWindowActivated = await ActivateBackgroundWindow("TallyPrime");
        if (!isWindowActivated)
        {
            return;
        }

        // Proceed with the entry
        RunEntry(exceldata, file);
    }
    #endregion

    #region Getting Tally Prime
    private async Task<bool> ActivateBackgroundWindow(string windowTitle)
    {
        // Find the window handle by its title
        IntPtr hWnd = Window_Handler.FindWindow(null, windowTitle);

        if (hWnd != IntPtr.Zero)
        {
            // Set the window to the foreground
            Window_Handler.SetForegroundWindow(hWnd);
            return true;
        }
        else
        {
            Window_Handler.Restore(App.MainWindow);
            await ShowDialog.ShowMsgBox("Error", $"Window '{windowTitle}' not found. \nPlease run '{windowTitle}'. ", "OK", null, 1, App.MainWindow);
            return false;
        }
    }
    #endregion

    #region FilePicker
    private async Task<StorageFile?> Filepick()
    {
        // Open File Picker to select a single Excel file
        var picker = new Windows.Storage.Pickers.FileOpenPicker();
        picker.FileTypeFilter.Add(".xlsx");
        picker.FileTypeFilter.Add(".xls");
        picker.SuggestedStartLocation = Windows.Storage.Pickers.PickerLocationId.DocumentsLibrary;

        var hwnd = WinRT.Interop.WindowNative.GetWindowHandle(App.MainWindow);
        WinRT.Interop.InitializeWithWindow.Initialize(picker, hwnd);

        var file = await picker.PickSingleFileAsync();

        if (file == null)
        {
            await ShowDialog.ShowMsgBox("Warning", "No Excel files were selected.", "OK", null, 1, App.MainWindow);
            return null;
        }
        return file;
    }
    #endregion

    #region Read Data from Excel
    private async Task<IEnumerable<VoucherCorrectionData>?> ReadData(StorageFile file)
    {
        var exceldata = new List<VoucherCorrectionData>(); // Initialize the exceldata list

        using (var workbook = new XLWorkbook(file.Path))
        {
            // Check if the "Ready Data" sheet exists  
            if (!workbook.TryGetWorksheet("Ready Data", out var worksheet))
            {
                await ShowDialog.ShowMsgBox("Warning", $"The 'Ready Data' sheet was not found in {file.Name}.", "OK", null, 1, App.MainWindow);
                return null;
            }

            // Get the headers as before (no changes here)
            var headers = new List<string> { "Voucher No." };

            // Map headers to column indices
            var existingHeaders = new Dictionary<string, int>();
            var firstRow = worksheet.Row(1);
            int lastColumn = worksheet.LastColumnUsed().ColumnNumber();

            for (int col = 1; col <= lastColumn; col++)
            {
                var headerCell = firstRow.Cell(col).GetString();
                if (!string.IsNullOrEmpty(headerCell))
                {
                    existingHeaders[headerCell] = col;
                }
            }

            if (!existingHeaders.ContainsKey("Voucher No."))
            {
                await ShowDialog.ShowMsgBox("The 'Voucher No.' header is missing.", "Warning", "OK", null, 1, App.MainWindow);
                return null;
            }

            // Iterate through rows starting from row 2
            var rows = worksheet.RowsUsed().Skip(1); // Skip header row
            foreach (var row in rows)
            {
                bool isRowEmpty = true;
                foreach (var cell in row.Cells(1, worksheet.LastColumnUsed().ColumnNumber()))
                {
                    if (!string.IsNullOrWhiteSpace(cell.GetString()))
                    {
                        isRowEmpty = false;
                        break;
                    }
                }
                if (isRowEmpty) continue;

                var entry = new VoucherCorrectionData
                {
                    Row = row.RowNumber(),
                    //VouchNo = string.IsNullOrWhiteSpace(row.Cell(existingHeaders["Voucher No."]).GetString()) ? "0" : row.Cell(existingHeaders["Voucher No."]).GetString()
                    VouchNo = string.IsNullOrWhiteSpace(row.Cell(existingHeaders["Voucher No."]).GetString()) ? "0" : row.Cell(existingHeaders["Voucher No."]).GetString()
                };

                exceldata.Add(entry);
            }
        }
        return exceldata;
    }

    #endregion

    #region Run Data Entry
    private async void RunEntry(IEnumerable<VoucherCorrectionData> exceldata, StorageFile file)
    {
        IEnumerable<VoucherCorrectionData> dataToProcess;
        dataToProcess = exceldata; // Process all skipped data

        await Task.Delay(TimeSpan.FromSeconds(1));

        foreach (var entry in dataToProcess)
        {
            keyboard.Enter_key();
            if (entry.VouchNo == "0")
            {
                keyboard.Esc();
                await Task.Delay(TimeSpan.FromSeconds(2));
            }
            else
            {
                keyboard.Input.TextEntry(entry.VouchNo);
            }

            keyboard.Ctrl_A();
            await Task.Delay(TimeSpan.FromSeconds(2));
            keyboard.Down();

        }
        // Final delay before restoring window
        await Task.Delay(TimeSpan.FromSeconds(2));
        Window_Handler.Restore(App.MainWindow);
        await ShowDialog.ShowMsgBox("Success", "All entries have been modified", "OK", null, 1, App.MainWindow);

    }
    #endregion

    public class VoucherCorrectionData
    {
        public int Row { get; set; }
        public string VouchNo { get; set; }
    }
}