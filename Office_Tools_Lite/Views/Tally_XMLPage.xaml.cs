using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Office_Tools_Lite.CLOSED_XML.XML_Generator;
using Office_Tools_Lite.Task_Helper;
using Office_Tools_Lite.ViewModels;

namespace Office_Tools_Lite.Views;

public sealed partial class Tally_XMLPage : Page
{
    public Tally_XMLViewModel ViewModel
    {
        get;
    }

    public Tally_XMLPage()
    {
        ViewModel = App.GetService<Tally_XMLViewModel>();
        InitializeComponent();
    }

    private async void Run_XML_Generator(object sender, RoutedEventArgs e)
    {
        if (sender is Button button)
        {
            Btn1.IsHitTestVisible = false;
            Btn2.IsHitTestVisible = false;
            Btn3.IsHitTestVisible = false;
            Btn4.IsHitTestVisible = false;
            Btn5.IsHitTestVisible = false;
            Btn6.IsHitTestVisible = false;
            Btn7.IsHitTestVisible = false;
            Btn8.IsHitTestVisible = false;
            Btn9.IsHitTestVisible = false;

            // Show the processing text
            ProcessingText.Visibility = Visibility.Visible;
            // Action mapping for buttons
            var actions = new Dictionary<string, Func<Task>>
        {
            { "Btn1", () => new Sales_Excel_To_Xml_Converter().Execute(App.MainWindow) },
            { "Btn2", () => new Sales_Return_Excel_To_Xml_Converter().Execute(App.MainWindow) },
            { "Btn3", () => new Purchase_Excel_To_Xml_Converter().Execute(App.MainWindow) },
            { "Btn4", () => new Purchase_Return_Excel_To_Xml_Convertor().Execute(App.MainWindow) },
            { "Btn5", () => new Bank_Excel_To_Xml_Converter().Execute(App.MainWindow) },
            { "Btn6", () => new Voucher_Excel_To_Xml_Converter().Execute(App.MainWindow) },
            { "Btn7", () => new Master_Excel_To_Xml_Converter_2().Execute(App.MainWindow) },
            { "Btn8", () => new Master_Excel_To_Xml_Converter_6().Execute(App.MainWindow) },
            { "Btn9", async () => await ShowDialog.ShowMsgBoxAsync("Almost there!", "This feature is part of the Full version.\n\nUpgrade now and enjoy everything without limits!","Keep using Lite","Upgrade me Direct!","Go to Web", 2,App.MainWindow) },

        };

            try
            {
                // Execute the appropriate action
                if (actions.TryGetValue(button.Name, out var action))
                {
                    await action();
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

                // Hide the processing text
                ProcessingText.Visibility = Visibility.Collapsed;
            }
        }
    }

    private void ShowAllBtn_Click(object sender, RoutedEventArgs e)
    {
        Btn1.Visibility = Visibility.Visible;
        Btn2.Visibility = Visibility.Visible;
        Btn3.Visibility = Visibility.Visible;
        Btn4.Visibility = Visibility.Visible;
        Btn5.Visibility = Visibility.Visible;
        Btn6.Visibility = Visibility.Visible;
        Btn7.Visibility = Visibility.Visible;
        Btn8.Visibility = Visibility.Visible;
        Btn9.Visibility = Visibility.Visible;

    }

    //private void Run_XML_Generators(object sender, RoutedEventArgs e)
    //{
    //    if (sender is Button button)
    //    {

    //        var actions = new Dictionary<string, Action>
    //    {
    //        { "Btn9", () => new File_Generator_run().Execute() },
    //        { "Btn10", () => new File_Generator_run().Execute() },
    //        { "Btn11", () => new File_Generator_run().Execute() },
    //        { "Btn9", () => this.Frame.Navigate(typeof(Voucher_Generator))},
    //        { "Btn10", () => this.Frame.Navigate(typeof(Turnover_Generator))},

    //    };

    //        // Execute the appropriate action
    //        if (actions.TryGetValue(button.Name, out var action))
    //        {
    //            action();
    //        }

    //    }

    //}
}
