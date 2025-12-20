using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Office_Tools_Lite.CLOSED_XML.Adjustments;
using Office_Tools_Lite.ViewModels;

namespace Office_Tools_Lite.Views;

public sealed partial class Sales_AdjustmentsPage : Page
{
    public Sales_AdjustmentsViewModel ViewModel
    {
        get;
    }

    public Sales_AdjustmentsPage()
    {
        ViewModel = App.GetService<Sales_AdjustmentsViewModel>();
        InitializeComponent();
    }

    private async void Run_Sales_Adjutments(object sender, RoutedEventArgs e)
    {
        if (sender is Button button)
        {
            // Disable all buttons
            // Make all buttons non-clickable while preserving appearance
            Btn1.IsHitTestVisible = false;
            Btn2.IsHitTestVisible = false;
            Btn3.IsHitTestVisible = false;


            ProcessingText.Visibility = Visibility.Visible;
            // Action mapping for buttons
            var actions = new Dictionary<string, Func<Task>>
            {
                { "Btn1", () => new Sales_Adjust_Self().Execute(App.MainWindow) },
                { "Btn2", () => new Sales_Adjust_others_X().Execute(App.MainWindow) },
                { "Btn3", () => new Sales_Adjust_others_S().Execute(App.MainWindow) },


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

                // Hide the processing text
                ProcessingText.Visibility = Visibility.Collapsed;
            }
        }
    }

}