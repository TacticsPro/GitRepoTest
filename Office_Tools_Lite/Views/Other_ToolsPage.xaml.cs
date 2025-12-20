using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Office_Tools_Lite.CLOSED_XML.Other_Tools;
using Office_Tools_Lite.Task_Helper;
using Office_Tools_Lite.ViewModels;

namespace Office_Tools_Lite.Views;

public sealed partial class Other_ToolsPage : Page
{
    public Other_ToolsViewModel ViewModel
    {
        get;
    }

    public Other_ToolsPage()
    {
        ViewModel = App.GetService<Other_ToolsViewModel>();
        InitializeComponent();
    }

    private async void Run_Other_Tools(object sender, RoutedEventArgs e)
    {
        if (sender is Button button)
        {
           
            var actions = new Dictionary<string, Action>
        {
            { "Btn1", () => this.Frame.Navigate(typeof(Excel_to_PDF)) },
            { "Btn2", () => this.Frame.Navigate(typeof(Notice_Letter_for_not_Uploaded)) },
            { "Btn3", () => this.Frame.Navigate(typeof(Notice_Letter_to_Tax_Office)) },
            { "Btn4", () => this.Frame.Navigate(typeof(Excel_Merger))},
            { "Btn5", () => this.Frame.Navigate(typeof(GSTR_3B_to_Excel))},
            { "Btn6", async () => await ShowDialog.ShowMsgBoxAsync("Almost there!", "This feature is part of the Full version.\n\nUpgrade now and enjoy everything without limits!","Keep using Lite","Upgrade me Direct!","Go to Web", 2,App.MainWindow) },
            { "Btn7", async () => await ShowDialog.ShowMsgBoxAsync("Almost there!", "This feature is part of the Full version.\n\nUpgrade now and enjoy everything without limits!","Keep using Lite","Upgrade me Direct!","Go to Web", 2,App.MainWindow) },
            { "Btn8", async () => await ShowDialog.ShowMsgBoxAsync("Almost there!", "This feature is part of the Full version.\n\nUpgrade now and enjoy everything without limits!","Keep using Lite","Upgrade me Direct!","Go to Web", 2,App.MainWindow) },
            { "Btn9", async () => await ShowDialog.ShowMsgBoxAsync("Almost there!", "This feature is part of the Full version.\n\nUpgrade now and enjoy everything without limits!","Keep using Lite","Upgrade me Direct!","Go to Web", 2,App.MainWindow) },
            { "Btn10", async () => await ShowDialog.ShowMsgBoxAsync("Almost there!", "This feature is part of the Full version.\n\nUpgrade now and enjoy everything without limits!","Keep using Lite","Upgrade me Direct!","Go to Web", 2,App.MainWindow) },
            { "Btn11", async () => await ShowDialog.ShowMsgBoxAsync("Almost there!", "This feature is part of the Full version.\n\nUpgrade now and enjoy everything without limits!","Keep using Lite","Upgrade me Direct!","Go to Web", 2,App.MainWindow) },
            { "Btn12", async () => await ShowDialog.ShowMsgBoxAsync("Almost there!", "This feature is part of the Full version.\n\nUpgrade now and enjoy everything without limits!","Keep using Lite","Upgrade me Direct!","Go to Web", 2,App.MainWindow) },
            { "Btn13", async () => await ShowDialog.ShowMsgBoxAsync("Almost there!", "This feature is part of the Full version.\n\nUpgrade now and enjoy everything without limits!","Keep using Lite","Upgrade me Direct!","Go to Web", 2,App.MainWindow) },

        };

            // Execute the appropriate action
            if (actions.TryGetValue(button.Name, out var action))
            {
                action();
            }

        }
       
    }

    //private void ShowAllBtn_Click(object sender, RoutedEventArgs e)
    //{
    //    Btn1.Visibility = Visibility.Visible;
    //    Btn2.Visibility = Visibility.Visible;
    //    Btn3.Visibility = Visibility.Visible;
    //    Btn4.Visibility = Visibility.Visible;
    //    Btn5.Visibility = Visibility.Visible;
    //    Btn6.Visibility = Visibility.Visible;
    //    Btn7.Visibility = Visibility.Visible;
    //    Btn8.Visibility = Visibility.Visible;
    //    Btn9.Visibility = Visibility.Visible;
    //    Btn10.Visibility = Visibility.Visible;
    //    Btn11.Visibility = Visibility.Visible;
    //    Btn12.Visibility = Visibility.Visible;
    //    Btn13.Visibility = Visibility.Visible;

    //}
}
