using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Office_Tools_Lite.CLOSED_XML.Adjustments;
using Office_Tools_Lite.CLOSED_XML.Data_Entry;
using Office_Tools_Lite.CLOSED_XML.XML_Generator;
using Office_Tools_Lite.Task_Helper;
using Office_Tools_Lite.ViewModels;

namespace Office_Tools_Lite.Views;

public sealed partial class Data_EntriesPage : Page
{
    public Data_EntriesViewModel ViewModel
    {
        get;
    }

    public Data_EntriesPage()
    {
        ViewModel = App.GetService<Data_EntriesViewModel>();
        InitializeComponent();
    }
    
    private async void Run_Data_Entries(object sender, RoutedEventArgs e)
    {

        if (sender is Button button)
        {
            // Disable all buttons
            // Make all buttons non-clickable while preserving appearance
            Btn1.IsHitTestVisible = false;
            Btn2.IsHitTestVisible = false;
            Btn3.IsHitTestVisible = false;
            Btn4.IsHitTestVisible = false;
            Btn5.IsHitTestVisible = false;
            Btn6.IsHitTestVisible = false;
            Btn7.IsHitTestVisible = false;
            Btn8.IsHitTestVisible = false;
            Btn9.IsHitTestVisible = false;
            Btn10.IsHitTestVisible = false;
            Btn11.IsHitTestVisible = false;

            var actions = new Dictionary<string, Action>
            {
                { "Btn1", async () => await ShowDialog.ShowMsgBoxAsync("Almost there!", "This feature is part of the Full version.\n\nUpgrade now and enjoy everything without limits!","Keep using Lite","Upgrade me Direct!","Go to Web", 2,App.MainWindow) },
                { "Btn2", async () => await ShowDialog.ShowMsgBoxAsync("Almost there!", "This feature is part of the Full version.\n\nUpgrade now and enjoy everything without limits!","Keep using Lite","Upgrade me Direct!","Go to Web", 2,App.MainWindow) },
                { "Btn3", async () => await ShowDialog.ShowMsgBoxAsync("Almost there!", "This feature is part of the Full version.\n\nUpgrade now and enjoy everything without limits!","Keep using Lite","Upgrade me Direct!","Go to Web", 2,App.MainWindow) },
                { "Btn4", async () => await ShowDialog.ShowMsgBoxAsync("Almost there!", "This feature is part of the Full version.\n\nUpgrade now and enjoy everything without limits!","Keep using Lite","Upgrade me Direct!","Go to Web", 2,App.MainWindow) },
                { "Btn5", async () => await ShowDialog.ShowMsgBoxAsync("Almost there!", "This feature is part of the Full version.\n\nUpgrade now and enjoy everything without limits!","Keep using Lite","Upgrade me Direct!","Go to Web", 2,App.MainWindow) },
                { "Btn6", async () => await ShowDialog.ShowMsgBoxAsync("Almost there!", "This feature is part of the Full version.\n\nUpgrade now and enjoy everything without limits!","Keep using Lite","Upgrade me Direct!","Go to Web", 2,App.MainWindow) },
                { "Btn7", async () => await ShowDialog.ShowMsgBoxAsync("Almost there!", "This feature is part of the Full version.\n\nUpgrade now and enjoy everything without limits!","Keep using Lite","Upgrade me Direct!","Go to Web", 2,App.MainWindow) },
                { "Btn8", async () => await ShowDialog.ShowMsgBoxAsync("Almost there!", "This feature is part of the Full version.\n\nUpgrade now and enjoy everything without limits!","Keep using Lite","Upgrade me Direct!","Go to Web", 2,App.MainWindow) },
                { "Btn9", async () => await ShowDialog.ShowMsgBoxAsync("Almost there!", "This feature is part of the Full version.\n\nUpgrade now and enjoy everything without limits!","Keep using Lite","Upgrade me Direct!","Go to Web", 2,App.MainWindow) },

            };

            try
            {
                // Execute the appropriate action
                if (actions.TryGetValue(button.Name, out var action))
                {
                    action();
                }
            }
            finally
            {
                // Re-enable clicking for all buttons
                Btn1.IsHitTestVisible = true;
                Btn2.IsHitTestVisible = true;
                Btn3.IsHitTestVisible = true;
                Btn4.IsHitTestVisible = true;
                Btn5.IsHitTestVisible = true;
                Btn6.IsHitTestVisible = true;
                Btn7.IsHitTestVisible = true;
                Btn8.IsHitTestVisible = true;
                Btn9.IsHitTestVisible = true;
                Btn10.IsHitTestVisible = true;
                Btn11.IsHitTestVisible = true;
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

    //}

    private async void Button_Click_10(object sender, RoutedEventArgs e)
    {
        var NarrationCorrection = new Narration_Correction();
        await NarrationCorrection.Execute();

    }

    private async void Button_Click_11(object sender, RoutedEventArgs e)
    {
        var voucherNoCorrection = new Voucher_No_Correction();
        await voucherNoCorrection.Execute();

    }

}