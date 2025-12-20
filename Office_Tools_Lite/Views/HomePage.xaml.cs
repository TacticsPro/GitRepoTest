using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Office_Tools_Lite.Task_Helper;
using Office_Tools_Lite.ViewModels;

namespace Office_Tools_Lite.Views;

public sealed partial class HomePage : Page
{
    public HomeViewModel ViewModel
    {
        get;
    }

    public HomePage()
    {
        ViewModel = App.GetService<HomeViewModel>();
        InitializeComponent();
        getFullVersion();

    }

    private void getFullVersion()
    {

        var releasePageUrl_lite = "https://github.com/TacticsPro/Office_Tools_Lite_Releases";
        hyperbutton1.NavigateUri = new Uri(releasePageUrl_lite);
        var releasePageUrl_full = FinderService.GetReleasePageUrl();
        hyperbutton2.NavigateUri = new Uri(releasePageUrl_full);
        directupdate.Click += async (s, e) =>
        {
            await ShowDialog.ShowMsgBoxAsync("Unlock Full Features","Upgrade to the full version for unlimited processing, advanced tools, and seamless workflows—no more limits on your productivity!",null,"Upgrade Now",null,2,App.MainWindow);
        };

    }

    private void pageopener(object sender, RoutedEventArgs e)
    {
        if (sender is Button button)
        {
            // Action mapping for buttons
            var actions = new Dictionary<string, Type>
        {
            { "Btn1", typeof(Purchase_AdjustmentsPage) },
            { "Btn2", typeof(CN_DN_AdjustmentsPage) },
            { "Btn3", typeof(Sales_AdjustmentsPage) },
            { "Btn4", typeof(GSTR1_AdjustmentsPage) },
            { "Btn5", typeof(Other_ToolsPage) },
            { "Btn6", typeof(Data_EntriesPage) },
            { "Btn7", typeof(Tally_XMLPage) }
        };

            // Navigate to the appropriate page
            if (actions.TryGetValue(button.Name, out var pageType))
            {
                this.Frame.Navigate(pageType);
            }
        }
    }


}