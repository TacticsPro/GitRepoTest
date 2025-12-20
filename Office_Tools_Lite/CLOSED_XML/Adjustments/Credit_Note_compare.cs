using System.Data;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Office_Tools_Lite.Task_Helper;
using Windows.Storage.Pickers;
using WinRT.Interop;

namespace Office_Tools_Lite.CLOSED_XML.Adjustments;

public class Credit_Note_compare
{
    #region Execute Main operation
    public async Task Execute(Window mainWindow)
    {
        await ShowDialog.ShowMsgBox("Notice", "Lite version can only process max 15 row data.\nYou can upgrade from Home Page for Full Feature!\n\nPress OK to Continue with Lite.", "OK", null, 1, App.MainWindow);

        #region Pick-up Credit Note File
        var dialog = await ShowDialog.ShowMsgBox(
            "Credit Note (Compare)",
            "This will merge all the Excel files (Credit Note Sheet) you select.\n\n Click OK to continue.",
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
            await ShowDialog.ShowMsgBox("Warning", "No Excel files were selected","Ok",null, 1, App.MainWindow);
            return;
        }
        #endregion

        #region Pick-up Creditors File
        var listFileDialog = await ShowDialog.ShowMsgBox(
            "Creditors List",
            "Pick the Excel file that contains the Creditors sheet.",
            "OK", "Cancel", 1, mainWindow);

        if (listFileDialog != ContentDialogResult.Primary)
        {
            return;
        }

        var filePicker1 = new FileOpenPicker();
        var hwnd1 = WindowNative.GetWindowHandle(mainWindow);
        InitializeWithWindow.Initialize(filePicker1, hwnd1);
        filePicker1.FileTypeFilter.Add(".xlsx");

        var listFile = await filePicker1.PickSingleFileAsync();
        if (listFile == null)
        {
            await ShowDialog.ShowMsgBox("Warning", "No Creditors file selected","Ok", null, 1, App.MainWindow);
            return;
        }
        #endregion

        var inputExcelFiles = files.Select(f => f.Path).ToArray();
        var originalDir = Path.GetDirectoryName(inputExcelFiles[0]);
        var subDir = Path.Combine(originalDir, "Output_Files");
        Directory.CreateDirectory(subDir);

        var inputSheetName = "Credit Note";
        var outputExcelFile = Path.Combine(subDir, "Credit_Note_compare.xlsx");
        var outputSheetName = "Credit Note";
        var FirstRowContent = "Date";
        var listFilePath = listFile.Path;
        var sourceSheetName = "Creditors";
        var targetSheetName = "Creditors";

        bool appendSuccess = await Excel_Appender.AppendExcelFiles(inputExcelFiles, inputSheetName, outputExcelFile, outputSheetName, FirstRowContent);

        if (appendSuccess)
        {
            bool addlistsheet = await Excel_Appender.AddListSheet(targetExcelFile: outputExcelFile, sourceExcelFile: listFilePath, sourceSheetName, targetSheetName);
            if (addlistsheet == false)
            {
                return;
            }
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
            byte[] decryptedContentBytes = Transformation.GetTransformedFileContent("Credit_Note_compare.dll");
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
            process.StartInfo.Arguments = $"\"{outputExcelFile}\" \"Credit_Note_compare\"";
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
