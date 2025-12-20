using System.Diagnostics;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Office_Tools_Lite.Task_Helper;
using Windows.Storage.Pickers;
using WinRT.Interop;

namespace Office_Tools_Lite.CLOSED_XML.Other_Tools;

public sealed partial class Excel_to_PDF : Microsoft.UI.Xaml.Controls.Page
{
    public Excel_to_PDF()
    {
        InitializeComponent();
    }

    private async void ProceedButton_Click(object sender, RoutedEventArgs e)
    {
        ProcessingText.Visibility = Visibility.Visible;

        try
        {
            // Get selected options
            var orientation = ((ComboBoxItem)OrientationComboBox.SelectedItem)?.Content.ToString();
            var paperSize = ((ComboBoxItem)PaperSizeComboBox.SelectedItem)?.Content.ToString();
            var zoomValue = int.TryParse(ZoomTextBox.Text, out int zoom) ? zoom : 100;
            var pageNumber = int.TryParse(PageNumberTextBox.Text, out int startPageNumber) ? startPageNumber : 1;
            var pageNumberType = ((ComboBoxItem)PageNumberTypeComboBox.SelectedItem)?.Content.ToString();

            // Open File Picker to select Excel files
            var picker = new FileOpenPicker();
            picker.FileTypeFilter.Add(".xlsx");
            picker.FileTypeFilter.Add(".xls");
            picker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;

            var hwnd = WindowNative.GetWindowHandle(App.MainWindow);
            InitializeWithWindow.Initialize(picker, hwnd);

            var files = await picker.PickMultipleFilesAsync();
            if (files == null || !files.Any())
            {
                ProcessingText.Visibility = Visibility.Collapsed;
                return;
            }

            // Prepare output directory
            var excelFiles = files.Select(f => f.Path).ToArray();
            var originalDir = Path.GetDirectoryName(excelFiles[0]);
            var outputDir = Path.Combine(originalDir, "Converted_pdf_Files");
            Directory.CreateDirectory(outputDir);

            // Prepare path to Excel_pdf.exe
            string macroRunnerDir = Path.Combine(AppContext.BaseDirectory, "Runner");
            string consoleExePath = Path.Combine(macroRunnerDir, "Excel_pdf.exe");

            if (!File.Exists(consoleExePath))
            {
                await ShowDialog.ShowMsgBox("Error", $"Excel_pdf.exe not found at: {consoleExePath}", "OK", null, 1, App.MainWindow);
                ProcessingText.Visibility = Visibility.Collapsed;
                return;
            }

            // Prepare arguments for console application
            string filePaths = string.Join(";", excelFiles);

            // Start the console application
            var process = new Process();
            process.StartInfo.FileName = consoleExePath;
            process.StartInfo.Arguments = $"\"{filePaths}\" \"{outputDir}\" \"{orientation}\" \"{paperSize}\" \"{zoomValue}\" \"{startPageNumber}\" \"{pageNumberType}\"";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.RedirectStandardError = true;
            process.StartInfo.CreateNoWindow = true;

            process.Start();
            string output = await process.StandardOutput.ReadToEndAsync();
            string error = await process.StandardError.ReadToEndAsync();
            await process.WaitForExitAsync();

            if (process.ExitCode == 0 && output.Contains("Success"))
            {
                await ShowDialog.ShowMsgBox("Success", "Excel files were successfully converted to PDF.", "OK", null, 1, App.MainWindow);
                Process.Start("explorer.exe", outputDir);
            }
            else
            {
                await ShowDialog.ShowMsgBox("Error", $"Conversion failed: {error}", "OK", null, 1, App.MainWindow);
            }
        }
        catch (Exception ex)
        {
            await ShowDialog.ShowMsgBox("Error", $"An error occurred: {ex.Message}", "OK", null, 1, App.MainWindow);
        }
        finally
        {
            ProcessingText.Visibility = Visibility.Collapsed;
        }
    }
}