using System.Reflection;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Animation;
using Office_Tools_Lite.Contracts.Services;
using Office_Tools_Lite.Helpers;
using Office_Tools_Lite.Task_Helper;
using Office_Tools_Lite.ViewModels;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.System;

namespace Office_Tools_Lite.Views;

public sealed partial class ShellPage : Page
{
    private CancellationTokenSource _downloadCancellationTokenSource;
    public ShellViewModel ViewModel {get;}
    public static ShellPage Instance {get; private set;}
    private readonly Check_for_Update updateChecker;
    private static readonly string firstTimeOpenFilePath = Path.Combine(Path.GetTempPath(),"firstTimeOpen.txt");
    private static readonly string updateCheckCountPath = Path.Combine(Path.GetTempPath(), "updateCheckCount.txt");
    private const int MAX_AUTO_UPDATE_CHECKS = 3;

    public ShellPage(ShellViewModel viewModel)
    {
        ViewModel = viewModel;
        Instance = this;
        InitializeComponent();
        SetVersionDetails();
        updateChecker = new Check_for_Update();
        this.Loaded += async (s, e) =>
        {
            await FinderService.InitializeAsync(App.MainWindow); // Ensure initialization
            await UpdateExpiryLabelAsync();
            await CheckForUpdatesOnLoadAsync(s, e);
            firsttimeopen();
        };
        ViewModel.NavigationService.Frame = NavigationFrame;
        ViewModel.NavigationViewService.Initialize(NavigationViewControl);

        App.MainWindow.ExtendsContentIntoTitleBar = true;
        App.MainWindow.SetTitleBar(AppTitleBar);
        App.MainWindow.Activated += MainWindow_Activated;
        AppTitleBarText.Text = "AppDisplayName".GetLocalized();
    }

    private async Task UpdateExpiryLabelAsync()
    {
        try
        {
            await FinderService.UpdateExpiryLabelAsync(ExpireDate);
        }
        catch (Exception ex)
        {
            ExpireDate.Text = "Error loading expiry details.";
            ExpireDate.Foreground = new SolidColorBrush(Microsoft.UI.Colors.Red);
        }
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        TitleBarHelper.UpdateTitleBar(RequestedTheme);
        KeyboardAccelerators.Add(BuildKeyboardAccelerator(VirtualKey.Left, VirtualKeyModifiers.Menu));
        KeyboardAccelerators.Add(BuildKeyboardAccelerator(VirtualKey.GoBack));
        var (emailId, _, _, _) = FinderService.GetCachedLicenseDetails();
        wlcmtxt.Text = $"Hello, {char.ToUpper(emailId[0]) + emailId.Split('@')[0][1..]} !";
    }

    private void MainWindow_Activated(object sender, WindowActivatedEventArgs args)
    {
        App.AppTitlebar = AppTitleBarText as UIElement;
    }

    private void NavigationViewControl_DisplayModeChanged(NavigationView sender, NavigationViewDisplayModeChangedEventArgs args)
    {
        AppTitleBar.Margin = new Thickness
        {
            Left = sender.CompactPaneLength * (sender.DisplayMode == NavigationViewDisplayMode.Minimal ? 2 : 1),
            Top = AppTitleBar.Margin.Top,
            Right = AppTitleBar.Margin.Right,
            Bottom = AppTitleBar.Margin.Bottom
        };
    }

    private static KeyboardAccelerator BuildKeyboardAccelerator(VirtualKey key, VirtualKeyModifiers? modifiers = null)
    {
        var keyboardAccelerator = new KeyboardAccelerator { Key = key };
        if (modifiers.HasValue)
        {
            keyboardAccelerator.Modifiers = modifiers.Value;
        }
        keyboardAccelerator.Invoked += OnKeyboardAcceleratorInvoked;
        return keyboardAccelerator;
    }

    private static void OnKeyboardAcceleratorInvoked(KeyboardAccelerator sender, KeyboardAcceleratorInvokedEventArgs args)
    {
        var navigationService = App.GetService<INavigationService>();
        args.Handled = navigationService.GoBack();
    }

    #region copy directory
    private void CopyDirectory(string sourceDir, string destDir)
    {
        Directory.CreateDirectory(destDir);

        // Copy all files
        foreach (var file in Directory.GetFiles(sourceDir))
        {
            string destFile = Path.Combine(destDir, Path.GetFileName(file));
            File.Copy(file, destFile, true);
        }

        // Copy all subdirectories
        foreach (var subDir in Directory.GetDirectories(sourceDir))
        {
            string destSubDir = Path.Combine(destDir, Path.GetFileName(subDir));
            CopyDirectory(subDir, destSubDir);
        }
    }
    #endregion

    #region Get Sample Files
    private async void GetSampleFiles_Click(object sender, RoutedEventArgs e)
    {
        var folderPicker = new FolderPicker
        {
            SuggestedStartLocation = PickerLocationId.Desktop
        };
        folderPicker.FileTypeFilter.Add("*");
        var hwnd = WinRT.Interop.WindowNative.GetWindowHandle(App.MainWindow);
        WinRT.Interop.InitializeWithWindow.Initialize(folderPicker, hwnd);
        StorageFolder folder = await folderPicker.PickSingleFolderAsync();
        if (folder != null)
        {
            var sampleFilesPath = Path.Combine(AppContext.BaseDirectory, "Sample_Files");

           
            var outputFolder = await folder.CreateFolderAsync("Sample_Files", CreationCollisionOption.OpenIfExists);

            CopyDirectory(sampleFilesPath, outputFolder.Path);

            
            await ShowDialog.ShowMsgBox("Done", "All files have been extracted successfully!", "OK", null, 1, App.MainWindow);
            System.Diagnostics.Process.Start("explorer.exe", outputFolder.Path);
        }
    }
    #endregion

    private async void firsttimeopen()
    {
        try
        {
            string content = File.ReadAllText(firstTimeOpenFilePath);

            if (content.Trim().Equals("True", StringComparison.OrdinalIgnoreCase))
            {
                await ShowDialog.ShowMsgBox(
                "Welcome 🎉",
                "Thank you for installing Office Tools Lite!\n\n" +
                "Before you get started, please open Settings and take a quick screenshot of your license details. " +
                "Keeping this handy will help if you ever need support.", "Ok", null, 1,
                App.MainWindow);


                NavigationFrame.Navigate(typeof(SettingsPage));
                // If you don’t want to show again, reset it:
                File.WriteAllText(firstTimeOpenFilePath, "False");
            }
        }
        catch
        {

        }
    }

    private void SetVersionDetails()
    {
        var assembly = Assembly.GetExecutingAssembly();
        var version = assembly.GetName().Version;
        VersionTextBlock.Content = $"Version: {version}";
    }

    private void Click_on_Get_Help(object sender, RoutedEventArgs e)
    {
        NavigationFrame.Navigate(typeof(Guide_Page));
    }
    private void Click_on_Visual_Help(object sender, RoutedEventArgs e)
    {
        NavigationFrame.Navigate(typeof(Visual_Helps));

    }


    #region Check Update on loadasync
    // Call this method inside your Loaded event (you already have CheckForUpdatesOnLoadAsync)
    private async Task CheckForUpdatesOnLoadAsync(object sender, RoutedEventArgs e)
    {
        try
        {
           
            var (isUpdateAvailable, latestVersion, downloadUrl, releaseNotes) = await updateChecker.IsUpdateAvailableAsync(App.MainWindow);

            if (isUpdateAvailable)
            {
                // ONLY after 3+ ignored attempts → force the popup
                if (!UpdateChecked())
                {
                    CheckUpdateButton_Click(null, null);
                }

                // Increment the count every time we show the blinking update
                IncrementUpdateCheckCount();

                Update_btn_blinker();
                UpdateAvailable.Visibility = Visibility.Visible;
            }
        }
        catch (Exception ex)
        {
            // Silently fail or log - don't annoy user on startup
            System.Diagnostics.Debug.WriteLine($"Update check failed: {ex.Message}");
        }
    }

    // Check if we should show auto-update reminder (less than 3 times)
    private bool UpdateChecked()
    {
        try
        {
            if (!File.Exists(updateCheckCountPath))
                return true;

            if (int.TryParse(File.ReadAllText(updateCheckCountPath).Trim(), out int count))
                return count < MAX_AUTO_UPDATE_CHECKS;  // count < 3 → true
        }
        catch { }
        return true;
    }

    private void IncrementUpdateCheckCount()
    {
        try
        {
            int currentCount = 0;
            if (File.Exists(updateCheckCountPath))
                int.TryParse(File.ReadAllText(updateCheckCountPath).Trim(), out currentCount);

            currentCount++;
            File.WriteAllText(updateCheckCountPath, currentCount.ToString());
            if (currentCount > 6)
            {
                File.WriteAllText(updateCheckCountPath, "0");
            }
        }
        catch { }
    }

    #endregion

    #region Update Blink
    private void Update_btn_blinker()
    {
        var blinkAnimation = new DoubleAnimation
        {
            From = 1.0,
            To = 0.0,
            Duration = new Duration(TimeSpan.FromSeconds(2)),
            AutoReverse = true,
            RepeatBehavior = RepeatBehavior.Forever
        };
        var storyboard = new Storyboard();
        storyboard.Children.Add(blinkAnimation);
        Storyboard.SetTarget(blinkAnimation, UpdateAvailable);
        Storyboard.SetTargetProperty(blinkAnimation, "Opacity");
        VersionTextBlock.Margin = new Thickness(0, -25, 0, 0);
        storyboard.Begin();
    }
    #endregion

    #region Update on Load
    //private async void CheckForUpdatesOnLoad(object sender, RoutedEventArgs e)
    //{

    //    try
    //    {
    //        var (isUpdateAvailable, latestVersion, downloadUrl, releaseNotes) = await updateChecker.IsUpdateAvailableAsync(App.MainWindow);

    //        if (isUpdateAvailable)
    //        {
    //            Update_btn_blinker();
    //            UpdateAvailable.Visibility = Visibility.Visible;

    //            // Show update dialog
    //            var dialog = new ContentDialog
    //            {
    //                Title = "Update Available",
    //                Content = $"A new version ({latestVersion}) is available.\n\nRelease Notes:\n{releaseNotes}",
    //                CloseButtonText = "Cancel",
    //                PrimaryButtonText = "Update",
    //                DefaultButton = ContentDialogButton.Primary,
    //                XamlRoot = this.Content.XamlRoot
    //            };

    //            var result = await dialog.ShowAsync();
    //            if (result == ContentDialogResult.Primary)
    //            {
    //                var file = await DownloadFileAsync(new Uri(downloadUrl));
    //                if (file != null)
    //                {
    //                    InstallUpdate(file);
    //                }
    //            }
    //            else
    //            {
    //                Environment.Exit(0);
    //            }
    //        }

    //    }
    //    catch (Exception ex)
    //    {
    //        await ShowDialog.ShowMsgBox("Error", $"Error checking for updates: {ex.Message}",App.MainWindow);
    //        Environment.Exit(0);
    //    }
    //}
    #endregion

    #region Check for Update Button Click
    public async void CheckUpdateButton_Click(object sender, RoutedEventArgs e)
    {
        
        try
        {
            var (isUpdateAvailable, latestVersion, downloadUrl, releaseNotes) = await updateChecker.IsUpdateAvailableAsync(App.MainWindow);

            if (isUpdateAvailable)
            {
                // Show update dialog
                var dialog = await ShowDialog.ShowMsgBox(
                    "Update Available",
                    $"A new version ({latestVersion}) is available.\n\nRelease Notes:\n{releaseNotes}",
                    "Update", "Cancel", 1, App.MainWindow);

                if (dialog == ContentDialogResult.Primary)
                {
                    var file = await DownloadFileAsync(new Uri(downloadUrl));
                    if (file != null)
                    {
                        InstallUpdate(file);
                    }
                }
                //else
                //{
                //    Environment.Exit(0);

                //}
            }
            else
            {
                await ShowDialog.ShowMsgBox("No new version", "You are already using the latest version.", "OK", null, 1, App.MainWindow);
            }
        }
        catch (Exception ex)
        {
            await ShowDialog.ShowMsgBox("Error", $"Error checking for updates: {ex.Message}", "Ok", null, 1, App.MainWindow);
        }
    }
    #endregion

    #region Download New File
    private async Task<StorageFile> DownloadFileAsync(Uri fileUri)
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
                XamlRoot = this.XamlRoot
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
    private async void InstallUpdate(StorageFile file)
    {
        // Reset the counter when user manually checks
        File.WriteAllText(updateCheckCountPath, "0");

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
}
