using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Office_Tools_Lite.CLOSED_XML.Adjustments;
using Office_Tools_Lite.Task_Helper;
using Office_Tools_Lite.ViewModels;

namespace Office_Tools_Lite.Views;

public sealed partial class GSTR1_AdjustmentsPage : Page
{
    public GSTR1_AdjustmentsViewModel ViewModel
    {
        get;
    }

    public GSTR1_AdjustmentsPage()
    {
        ViewModel = App.GetService<GSTR1_AdjustmentsViewModel>();
        InitializeComponent();
    }

    private async void Run_HSN_Adjutments(object sender, RoutedEventArgs e)
    {
        if (sender is Button button)
        {
            // Disable all buttons
            // Make all buttons non-clickable while preserving appearance
            Btn1.IsHitTestVisible = false;
            Btn2.IsHitTestVisible = false;

            // Show the processing text
            ProcessingText.Visibility = Visibility.Visible;
            // Action mapping for buttons
            var actions = new Dictionary<string, Func<Task>>
            {
                { "Btn1", () => new HSN_json_direct().Execute(App.MainWindow) },
                { "Btn2", () => new HSN_json_with_adjust().Execute(App.MainWindow) },
                { "Btn3", () => new HSN_Descriptions_Fill().Execute(App.MainWindow) },
                { "Btn4", () => new GSTR1().Execute() },
                { "Btn5", () => new HSN_GSTR1_error_check().Execute(App.MainWindow) },

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

                // Hide the processing text
                ProcessingText.Visibility = Visibility.Collapsed;
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


    //}
}
