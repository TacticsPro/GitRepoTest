using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Windows.Storage;
using Windows.System;

namespace Office_Tools_Lite.Task_Helper;

public static class ShowDialog
{
    private static CancellationTokenSource _downloadCancellationTokenSource;

    public static async Task<ContentDialogResult> ShowMsgBox(string title, object content,
    string button1 = null, string button2 = null, decimal selectedButton = 0, Window mainWindow = null)
    {
        var dialog = new ContentDialog
        {
            Title = title,
            Content = content,
            PrimaryButtonText = button1,
            SecondaryButtonText = button2,
            DefaultButton = selectedButton switch
            {
                1 => ContentDialogButton.Primary,
                2 => ContentDialogButton.Secondary,
                _ => ContentDialogButton.None
            },

            XamlRoot = mainWindow.Content.XamlRoot
        };

        return await dialog.ShowAsync();
    }

    public static async Task<ContentDialogResult> ShowMsgBoxAsync(
    string title,
    object content,
    string primaryButtonText = null,
    string secondaryButtonText = null,
    string closeButtonText = null,
    decimal defaultButtonIndex = 0,
    Window mainWindow = null)
    {
        var releaseUrl = FinderService.GetReleasePageUrl();

        var dialog = new ContentDialog()
        {
            Title = title,
            Content = content,
            PrimaryButtonText = primaryButtonText,
            SecondaryButtonText = secondaryButtonText,
            CloseButtonText = closeButtonText,
            DefaultButton = defaultButtonIndex switch
            {
                1 => ContentDialogButton.Primary,
                2 => ContentDialogButton.Secondary,
                3 => ContentDialogButton.Close,
                _ => ContentDialogButton.None
            },
            XamlRoot = mainWindow?.Content.XamlRoot
        };

        if (!string.IsNullOrEmpty(closeButtonText) && Uri.TryCreate(releaseUrl, UriKind.Absolute, out var uri))
        {
            dialog.CloseButtonClick += async (s, e) =>
            {
                await Windows.System.Launcher.LaunchUriAsync(uri);
                dialog.Hide();
            };
        }

        dialog.SecondaryButtonClick += async (s, e) =>
        {
            Getfullversionbutton();
            dialog.Hide();
        };

        return await dialog.ShowAsync();
    }

    #region Check for Update Button Click
    private static async void Getfullversionbutton()
    {
        var updateChecker = new Check_for_Update();

        var downloadUrl = await updateChecker.getfullversion();
        var file = await DownloadFileAsync(new Uri(downloadUrl));
        if (file != null)
        {
            InstallUpdate(file);
        }
    }
    #endregion

    #region Download New File
    private static async Task<StorageFile> DownloadFileAsync(Uri fileUri)
    {
        _downloadCancellationTokenSource = new CancellationTokenSource();

        try
        {
            var httpClient = new HttpClient();
            var response = await httpClient.GetAsync(fileUri, HttpCompletionOption.ResponseHeadersRead, _downloadCancellationTokenSource.Token);
            response.EnsureSuccessStatusCode();

            var fileName = Path.GetFileName(fileUri.LocalPath);

            var downloadsFolder = await StorageFolder.GetFolderFromPathAsync(
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads"));

            var file = await downloadsFolder.CreateFileAsync(fileName, CreationCollisionOption.ReplaceExisting);

            var totalBytes = response.Content.Headers.ContentLength.GetValueOrDefault(0);
            var buffer = new byte[8192];
            var bytesRead = 0L;

            var progressDialog = new ContentDialog
            {
                Title = "Update",
                Content = "Download Progress: 0%",
                CloseButtonText = "Cancel",
                XamlRoot = App.MainWindow.Content.XamlRoot
            };

            // Handle cancel button in dialog
            progressDialog.CloseButtonClick += (_, _) =>
            {
                _downloadCancellationTokenSource.Cancel();
                progressDialog.Hide();
            };

            _ = progressDialog.ShowAsync();

            using (var inputStream = await response.Content.ReadAsStreamAsync())
            using (var outputStream = await file.OpenStreamForWriteAsync())
            {
                int read;
                while ((read = await inputStream.ReadAsync(buffer, 0, buffer.Length, _downloadCancellationTokenSource.Token)) != 0)
                {
                    await outputStream.WriteAsync(buffer, 0, read, _downloadCancellationTokenSource.Token);
                    bytesRead += read;

                    var progressPercentage = (int)((bytesRead * 100) / totalBytes);
                    progressDialog.Content = $"Download Progress: {progressPercentage}%";
                }
            }

            progressDialog.Hide();
            return file;
        }
        catch (OperationCanceledException)
        {
            await ShowDialog.ShowMsgBox("Download Canceled", "The update download was canceled.", "OK", null, 1, App.MainWindow);
            //Environment.Exit(0);
            return null;
        }
        catch (Exception ex)
        {
            await ShowDialog.ShowMsgBox("Error", $"Error downloading file: {ex.Message}", "Ok", null, 1, App.MainWindow);
            //Environment.Exit(0);
            return null;
        }
        finally
        {
            _downloadCancellationTokenSource = null;
        }
    }
    #endregion

    #region Install File
    private static async void InstallUpdate(StorageFile file)
    {

        try
        {
            var options = new LauncherOptions { TreatAsUntrusted = false };
            var success = await Launcher.LaunchFileAsync(file, options);

            if (!success)
            {
                await ShowDialog.ShowMsgBox("Error", "Failed to launch the update installer.", "OK", null, 1, App.MainWindow);
            }
        }
        catch (Exception ex)
        {
            await ShowDialog.ShowMsgBox("Error", $"Error installing update: {ex.Message}", "Ok", null, 1, App.MainWindow);
        }
    }
    #endregion

    public static async Task<string> ShowInputDialog(string title, Window mainWindow)
    {
        // Create a StackPanel to hold the content
        StackPanel panel = new StackPanel();

        // Create a Grid for labels and inputs
        Grid inputGrid = new Grid
        {
            Margin = new Thickness(0),
            ColumnSpacing = 10 // Add spacing between columns
        };

        // Define two rows in the grid
        inputGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        inputGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        // Define three columns for Financial Year and Month
        inputGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength( 1, GridUnitType.Auto) });
        inputGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Auto) });
        inputGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Auto) });

        // Financial Year label
        TextBlock financialYearLabel = new TextBlock
        {
            Text = "Financial Year:",
            Width = 100,
            Margin = new Thickness(10, 10, 0, 0)
        };
        Grid.SetRow(financialYearLabel, 0);
        Grid.SetColumn(financialYearLabel, 0);
        inputGrid.Children.Add(financialYearLabel);

        // Financial Year TextBox
        TextBox FYtextBox = new TextBox
        {
            Width = 120,
            Margin = new Thickness(0, 10, 0, 0)
        };
        Grid.SetRow(FYtextBox, 1);
        Grid.SetColumn(FYtextBox, 0);
        inputGrid.Children.Add(FYtextBox);

        // Month label
        TextBlock monthLabel = new TextBlock
        {
            Text = "Month:",
            Width = 100,
            Margin = new Thickness(25, 10, 0, 0)
        };
        Grid.SetRow(monthLabel, 0);
        Grid.SetColumn(monthLabel, 1);
        inputGrid.Children.Add(monthLabel);

        // Month ComboBox
        ComboBox monthComboBox = new ComboBox
        {
            Width = 100,
            Margin = new Thickness(10, 10, 0, 0)
        };
        string[] months = { "All", "Apr", "May", "June", "July", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar" };
        foreach (string month in months)
        {
            monthComboBox.Items.Add(month);
        }
        monthComboBox.SelectedIndex = 0; // Default to "All"
        Grid.SetRow(monthComboBox, 1);
        Grid.SetColumn(monthComboBox, 1);
        inputGrid.Children.Add(monthComboBox);

        // ledger label
        TextBlock ledgerLabel = new TextBlock
        {
            Text = "Create Party Ledger?",
            Margin = new Thickness(5, 10, 0, 0)
        };
        Grid.SetRow(ledgerLabel, 0);
        Grid.SetColumn(ledgerLabel, 2);
        inputGrid.Children.Add(ledgerLabel);

        // Month ComboBox
        ComboBox YesNo = new ComboBox
        {
            Width = 120,
            Margin = new Thickness(0, 10, 0, 0)
        };
        string[] yesno = { "Yes", "No" };
        foreach (string ans in yesno)
        {
            YesNo.Items.Add(ans);
        }
        YesNo.SelectedIndex = 0; // Default to "All"
        Grid.SetRow(YesNo, 1);
        Grid.SetColumn(YesNo, 2);
        inputGrid.Children.Add(YesNo);

        // Set Financial Year and Month defaults
        int currentMonth = DateTime.Now.Month;
        int currentYear = DateTime.Now.Year;
        int Year1, Year2;

        if (currentMonth > 4)
        {
            Year1 = currentYear;
            Year2 = (currentYear + 1) % 100;
        }
        else
        {
            Year1 = currentYear - 1;
            Year2 = currentYear % 100;
        }

        FYtextBox.Text = $"{Year1}-{Year2}";

        // Select Month in ComboBox
        int[] monthIndices = { 9, 10, 11, 12, 1, 2, 3, 4, 5, 6, 7, 8 };
        monthComboBox.SelectedIndex = monthIndices[currentMonth - 1];

        // Add controls to panel
        panel.Children.Add(inputGrid);

        // Create ContentDialog
        var dialog = await ShowDialog.ShowMsgBox(title, panel, "OK", "Cancel", 1, mainWindow);

        if (dialog == ContentDialogResult.Primary)
        {
            return $"{monthComboBox.SelectedItem?.ToString()}|{FYtextBox.Text}|{YesNo.SelectedItem?.ToString()}";
        }
        else
        {
            return string.Empty;
        }
    }
}