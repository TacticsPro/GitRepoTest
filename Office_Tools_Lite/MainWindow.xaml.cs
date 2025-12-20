using System.Diagnostics;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Media.Animation;
using Office_Tools_Lite.Helpers;
using Windows.UI.ViewManagement;

namespace Office_Tools_Lite;

public sealed partial class MainWindow : WindowEx
{
    private Microsoft.UI.Dispatching.DispatcherQueue dispatcherQueue;

    private UISettings settings;
    private Storyboard _blinkStoryboard;
    private CancellationTokenSource _downloadCancellationTokenSource;

    public MainWindow()
    {
        
        InitializeComponent();
        //_ = gettingfiles();
        //EnableVBA.Enable_All();
        //_ = RunEnableVBAAsync();

        AppWindow.SetIcon(Path.Combine(AppContext.BaseDirectory, "Assets/WindowIcon.ico"));
        Content = null;
        Title = "AppDisplayName".GetLocalized();

        // Theme change code picked from https://github.com/microsoft/WinUI-Gallery/pull/1239
        dispatcherQueue = Microsoft.UI.Dispatching.DispatcherQueue.GetForCurrentThread();
        settings = new UISettings();
        settings.ColorValuesChanged += Settings_ColorValuesChanged; // cannot use FrameworkElement.ActualThemeChanged event
        this.Closed += OnWindowClosed;
        
    }
    private void OnWindowClosed(object sender, WindowEventArgs args)
    {
        // Stop any animations or UI activity here
        _blinkStoryboard?.Stop();
        _blinkStoryboard = null;

        // Cancel and dispose of CancellationTokenSource
        _downloadCancellationTokenSource?.Cancel();
        _downloadCancellationTokenSource?.Dispose();
        _downloadCancellationTokenSource = null;

        // Terminate all Office_Tools_Lite processes (excluding the current process)
        int currentProcessId = Process.GetCurrentProcess().Id;
        foreach (var process in Process.GetProcessesByName("Office_Tools_Lite"))
        {
            try
            {
                if (process.Id != currentProcessId)
                {
                    process.Kill();
                    process.WaitForExit(1000);
                }
            }
            catch (Exception)
            {
                // Optionally log error
            }
        }

        // Note: No need to call Application.Current.Shutdown() in WinUI 3
    }

    private void Settings_ColorValuesChanged(UISettings sender, object args)
    {
        // This calls comes off-thread, hence we will need to dispatch it to current app's thread
        dispatcherQueue.TryEnqueue(() =>
        {
            TitleBarHelper.ApplySystemThemeToCaptionButtons();
        });
    }


}
