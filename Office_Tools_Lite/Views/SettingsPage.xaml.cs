using System.Diagnostics;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Office_Tools_Lite.Task_Helper;
using Office_Tools_Lite.ViewModels;

namespace Office_Tools_Lite.Views;

// TODO: Set the URL for your privacy policy by updating SettingsPage_PrivacyTermsLink.NavigateUri in Resources.resw.
public sealed partial class SettingsPage : Page
{
    public SettingsViewModel ViewModel
    {
        get;
    }

    public SettingsPage()
    {
        ViewModel = App.GetService<SettingsViewModel>();
        InitializeComponent();
        settingdetails();
    }

    private void settingdetails()
    {
        // Retrieve cached license details from LicenceService
        var (emailId, productId, machineId, activationTime) = FinderService.GetCachedLicenseDetails();
        EmailID.Text = $"Email ID : {emailId}";
        ProductID.Text = $"Product ID : {productId}";
        MachineID.Text = $"Machine ID : {machineId}";
        ActivationTime.Text = $"Activation Time : {activationTime}";

        var releasePageUrl = "https://github.com/TacticsPro/Office_Tools_Lite_Releases";
        hyperbutton.NavigateUri = new Uri(releasePageUrl);
    }
}
