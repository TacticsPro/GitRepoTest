using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Office_Tools_Lite.Task_Helper;
using Windows.System;

namespace Office_Tools_Lite.Views;

public sealed partial class Visual_Helps : Page
{
    private string? htmlFileUri;
    private static string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
    private static string targetTutorialsFolder = Path.Combine(localAppData, "Office_Tools_Lite", "Tutorials");
    public Visual_Helps()
    {
        this.InitializeComponent();
        pageloader();
    }

    private async void pageloader()
    {
        await onloadvisualHelpPage();
    }

    private async Task onloadvisualHelpPage()
    {
        // Ensure WebView2 is initialized
        await webview.EnsureCoreWebView2Async();

        // Disable the status bar to prevent URLs from showing on hover
        webview.CoreWebView2.Settings.IsStatusBarEnabled = false;

        htmlFileUri = App.GetCachedHtmlPath();

        //if (!Directory.Exists(targetTutorialsFolder))
        //{
        //    string sourceTutorialsFolder = Path.Combine(AppContext.BaseDirectory, "Tutorials");
        //    string targetFilePath = Path.Combine(sourceTutorialsFolder, "index.html");
        //    htmlFileUri = new Uri($"file:///{targetFilePath.Replace("\\", "/")}").AbsoluteUri;
        //    webview.Source = new Uri(htmlFileUri);
        //}
        //else
        //{
        //    // Load in WebView2
        //    webview.Source = new Uri(htmlFileUri);
        //}
        webview.Source = new Uri(htmlFileUri);
    }

    private async void LaunchBrowserButton_Click(object sender, RoutedEventArgs e)
    {
        if (!string.IsNullOrEmpty(htmlFileUri))
        {
            await Launcher.LaunchUriAsync(new Uri(htmlFileUri));
        }
    }
}
