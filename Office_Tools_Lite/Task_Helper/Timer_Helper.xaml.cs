using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Windows.Storage;

namespace Office_Tools_Lite.Task_Helper
{
    public sealed partial class TimerWindow : WindowEx
    {
        public TimerWindow()
        {
            this.InitializeComponent();
            SetWindowIcon();
        }

        // Method to start the countdown timer
        public async Task StartTimer(int seconds)
        {
            for (int i = seconds; i >= 0; i--)
            {
                TimerTextBlock.Text = $"Starting in: {i} Seconds"; // Update the timer on the UI
                await Task.Delay(1000); // Delay for 1 second
            }
            this.Close(); // Close the timer window after the countdown
        }

        private async void SetWindowIcon()
        {
            var appWindow = this.AppWindow;
            var iconPath = Path.Combine(AppContext.BaseDirectory, "Assets", "WindowIcon.ico");
            appWindow.SetIcon(iconPath);
        }
    }
}
