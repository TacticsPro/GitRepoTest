using System.Diagnostics;
using ClosedXML.Excel;
using Microsoft.UI.Xaml.Controls;
using Office_Tools_Lite.Task_Helper;
using Windows.Storage.Pickers;
using WinRT.Interop;

namespace Office_Tools_Lite.CLOSED_XML.Adjustments;

public class GSTR1
{
    #region Execute Main operation
    public async Task Execute()
    {
        await ShowDialog.ShowMsgBox("Notice", "Lite version can only process max 15 row data.\nYou can upgrade from Home Page for Full Feature!\n\nPress OK to Continue with Lite.", "OK", null, 1, App.MainWindow);

        var result =  await ShowDialog.ShowMsgBox("GSTR-1", "Select GSTR-1 Excel file", "OK", "Cancel", 1, App.MainWindow);
        if (result != ContentDialogResult.Primary)
        {
            return;
        }

        var filePicker = new FileOpenPicker();
        var hwnd = WindowNative.GetWindowHandle(App.MainWindow);
        InitializeWithWindow.Initialize(filePicker, hwnd);

        filePicker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
        filePicker.FileTypeFilter.Add(".xlsx");

        var files = await filePicker.PickMultipleFilesAsync();
        if (files.Count == 0)
        {
            await ShowDialog.ShowMsgBox("Warning","No Excel files were selected.","OK", null, 1, App.MainWindow);
            return;
        }

        var excelPath = files[0].Path;
        var originalDir = Path.GetDirectoryName(excelPath);

        using (var workbook = new XLWorkbook(excelPath))
        {
            var sheetName = "Ready Data";
            if (!workbook.Worksheets.Contains(sheetName))
            {
                await ShowDialog.ShowMsgBox("Error", "Worksheet 'Ready Data' not found in the Excel file.", "OK", null, 1, App.MainWindow);
                return;
            }
        }

        try
        {
            while (true) // Outer loop for GSTIN
            {
                var gstinTextBox = new TextBox { PlaceholderText = "Enter GSTIN" };

                var gstinDialog = await ShowDialog.ShowMsgBox(
                    "GSTIN Input", gstinTextBox,"OK", "Cancel", 1, App.MainWindow);

                if (gstinDialog != ContentDialogResult.Primary ||
                    string.IsNullOrWhiteSpace(gstinTextBox.Text))
                {
                    return;
                }

                var gstin = gstinTextBox.Text.Trim();

                while (true) // Inner loop for FP
                {
                    var financialPeriodTextBox = new TextBox
                    {
                        PlaceholderText = "Enter MMYYYY",
                        Text = $"{DateTime.Today.Month - 1:D2}{DateTime.Today.Year}"
                    };

                    var fpDialog = await ShowDialog.ShowMsgBox(
                        "Financial Period Input",
                        financialPeriodTextBox, "OK", "Back", 1, App.MainWindow);

                    if (fpDialog == ContentDialogResult.Secondary)
                        break; // Back to GSTIN

                    if (fpDialog != ContentDialogResult.Primary ||
                        string.IsNullOrWhiteSpace(financialPeriodTextBox.Text))
                    {
                        return;
                    }

                    var fp = financialPeriodTextBox.Text.Trim();
                    var outputJsonFile = Path.Combine(originalDir, $"GSTR1_{gstin}_{fp}.json");

                    // ✨ CALL THE NEW ENGINE HERE
                    await GSTR1_json_generator.ConvertExcelToGstr1Json(excelPath, outputJsonFile, gstin, fp);

                    

                    return; // Done
                }
            }
        }
        catch (Exception ex)
        {
           await ShowDialog.ShowMsgBox("Error", $"An error occurred: {ex.Message}", "OK", null, 1, App.MainWindow);

        }
    }

    #endregion
}
