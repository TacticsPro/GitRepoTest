using System.Data;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Office_Tools_Lite.Task_Helper;
using Windows.Storage.Pickers;
using WinRT.Interop;

namespace Office_Tools_Lite.CLOSED_XML.Adjustments;

public class Tally_Data_2B_Next_Oct_24
{
    #region Execute Main operation
    public async Task Execute(Window mainWindow)
    {
        await ShowDialog.ShowMsgBox("Notice", "Lite version can only process max 15 row data.\nYou can upgrade from Home Page for Full Feature!\n\nPress OK to Continue with Lite.", "OK", null, 1, App.MainWindow);

        var dialog = await ShowDialog.ShowMsgBox(
            "Tally Data 2B Next Oct-24",
            "To pick an Excel file, click OK to continue.",
            "OK", "Cancel", 1, mainWindow);

        if (dialog != ContentDialogResult.Primary)
        {
            return;
        }

        // File picker dialog
        var filePicker = new FileOpenPicker();
        var hwnd = WindowNative.GetWindowHandle(mainWindow);
        InitializeWithWindow.Initialize(filePicker, hwnd);

        filePicker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
        filePicker.FileTypeFilter.Add(".xlsx");

        var files = await filePicker.PickMultipleFilesAsync();
        if (files.Count == 0)
        {
            await ShowDialog.ShowMsgBox("Warning", "No Excel files were selected", "OK", null, 1, App.MainWindow);
            return;
        }

        var inputExcelFiles = files.Select(f => f.Path).ToArray();
        var originalDir = Path.GetDirectoryName(inputExcelFiles[0]);
        var subDir = Path.Combine(originalDir, "Output_Files");
        Directory.CreateDirectory(subDir);

        var inputSheetName = "Tally Data";
        var outputExcelFile = Path.Combine(subDir, "Tally(Adjusted_data).xlsx");
        var outputSheetName = "Tally Data";
        var FirstRowContent = "Date";

        bool appendSuccess = await Excel_Appender.AppendExcelFiles(inputExcelFiles, inputSheetName, outputExcelFile, outputSheetName, FirstRowContent);

        if (appendSuccess)
        {
            await RunMacro(mainWindow, outputExcelFile);
        }
    }
    #endregion

    #region  Run macro
    private async Task RunMacro(Window mainWindow, string outputExcelFile)
    {
        // Minimize the main application window
        //Window_Handler.Minimize(mainWindow);

        try
        {
            byte[] decryptedContentBytes = Transformation.GetTransformedFileContent("Tally_Data_2B_Next_Oct_24.dll");
            string vbContent = System.Text.Encoding.UTF8.GetString(decryptedContentBytes);

            string macroRunnerDir = Path.Combine(AppContext.BaseDirectory, "Runner");
            string helperExePath = Path.Combine(macroRunnerDir, "M_All.exe");

            if (!File.Exists(helperExePath))
            {
                await ShowDialog.ShowMsgBox("Error", $"M_All.exe not found at: {helperExePath}", "OK", null, 1, mainWindow);
                return;
            }

            var process = new System.Diagnostics.Process();
            process.StartInfo.FileName = helperExePath;
            process.StartInfo.Arguments = $"\"{outputExcelFile}\" \"Tally_Data_2B_Next_Oct_24\"";
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
                // Restore the main window
                //Window_Handler.Restore(mainWindow);
                await ShowDialog.ShowMsgBox("Success", "File Saved successfully.", "OK", null, 1, mainWindow);
                // Open the output folder in Explorer
                var outputFolderPath = System.IO.Path.GetDirectoryName(outputExcelFile);
                System.Diagnostics.Process.Start("explorer.exe", outputFolderPath);
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

}
