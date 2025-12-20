using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Office_Tools_Lite.CLOSED_XML.Adjustments;
using Office_Tools_Lite.CLOSED_XML.Other_Tools;
using Office_Tools_Lite.Task_Helper;
using Office_Tools_Lite.ViewModels;

namespace Office_Tools_Lite.Views;

public sealed partial class Purchase_AdjustmentsPage : Page
{
    public Purchase_AdjustmentsViewModel ViewModel
    {
        get;
    }

    public Purchase_AdjustmentsPage()
    {
        ViewModel = App.GetService<Purchase_AdjustmentsViewModel>();
        InitializeComponent();
    }

    private async void Run_Purchase_Adjutments(object sender, RoutedEventArgs e)
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

            // Show the processing text
            ProcessingText.Visibility = Visibility.Visible;
            // Action mapping for buttons
            var actions = new Dictionary<string, Func<Task>>
            {
                { "Btn1", () => new Purchase_csv_entry().Execute(App.MainWindow) },
                { "Btn2", () => new Purchase_csv_compare().Execute(App.MainWindow) },
                { "Btn3", () => new GSTR_2A_B2B_excel_compare().Execute(App.MainWindow) },
                { "Btn4", () => new GSTR_8A_B2B_compare_Upto_22_23().Execute(App.MainWindow) },
                { "Btn5", () => new GSTR_8A_B2B_compare_Next_23_24().Execute(App.MainWindow) },
                { "Btn6", () => new GSTR_2B_B2B_compare_Till_Sep_24().Execute(App.MainWindow) },
                { "Btn7", () => new GSTR_2B_B2B_compare_Next_Oct_24().Execute(App.MainWindow) },
                { "Btn8", () => new Tally_Data().Execute(App.MainWindow) },
                { "Btn9", () => new Tally_Data_2B_Next_Oct_24().Execute(App.MainWindow) },
                //{ "Btn10", () => new File_Generator_run().Execute() },

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
                Btn10.IsHitTestVisible = true;

                // Hide the processing text
                ProcessingText.Visibility = Visibility.Collapsed;
            }
        }
    }

    private async void Open_GSTR_Tally(object sender, RoutedEventArgs e)
    {

        if (sender is Button button)
        {

            var actions = new Dictionary<string, Action>
            {

                { "Btn10", async () => await ShowDialog.ShowMsgBox("Lite", "to use this you need Full veriosn!","OK","Cancel",1,App.MainWindow) },



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


    //}
}
