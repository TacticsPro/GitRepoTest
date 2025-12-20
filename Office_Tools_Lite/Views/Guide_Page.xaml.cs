using System.Diagnostics;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Office_Tools_Lite.Task_Helper;
using Office_Tools_Lite.Views;

namespace Office_Tools_Lite.Views;

public sealed partial class Guide_Page : Page
{
    public Guide_Page()
    {
        this.InitializeComponent();
        onload();
    }

    private void Click_on_GoBack(object sender, RoutedEventArgs e)
    {
        this.Frame.Navigate(typeof(HomePage));
    }
    private void onload()
    {
        var releasePageUrl = FinderService.GetReleasePageUrl();
        hyperbutton.NavigateUri = new Uri(releasePageUrl);
    }
}
