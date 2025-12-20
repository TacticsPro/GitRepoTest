using System.Threading.Tasks;
using Microsoft.UI;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;

namespace Office_Tools_Lite.Task_Helper;

public static class FinderService
{
    private static DateTime? _cachedExpiryDate;
    private static DateTime _cachedCurrentDate;
    private static string _cachedEmailId;
    private static string _cachedProductId;
    private static string _cachedMachineId;
    private static DateTime? _cachedActivationTime;
    private static string _cachedInfoPath;
    private static Finder _finderCheck;

    #region Initialize Finder Service
    public static async Task InitializeAsync(Window owner)
    {
        if (_cachedExpiryDate == null)
        {
            _finderCheck = new Finder();
            // 🔥 FETCH INTERNET DATE *BEFORE* LICENSE CHECK
            _cachedCurrentDate = await TimeService.GetInternetTimeAsync();

            // If Finder.InitializeDate hasn't finished, override with correct value
            _finderCheck.ForceSetCurrentDate(_cachedCurrentDate);

            _cachedExpiryDate = await _finderCheck.GetNewExpiryDate(owner);
            var (emailId, productId, machineId, activationTime) = await _finderCheck.GetLicenseDetails();
            _cachedEmailId = emailId;
            _cachedProductId = productId;
            _cachedMachineId = machineId;
            _cachedActivationTime = activationTime != default ? activationTime : null;

            try
            {
                var infoPath = await Finder.GetInfoPath();
                _cachedInfoPath = infoPath.info_path;

            }
            catch (Exception ex)
            {
                _cachedInfoPath = null; // Or set a default URL
            }
        }
    }
    #endregion

    #region Get Release Page
    public static string GetReleasePageUrl()
    {
        return _finderCheck?.GetReleasePage()
            ?? throw new InvalidOperationException("Licence_Check not initialized.");
    }
    #endregion

    #region Update Expiry Label
    public static async Task UpdateExpiryLabelAsync(TextBlock label)
    {
        try
        {
            if (_cachedExpiryDate == null)
            {
                label.Text = "Error: Expiry date not loaded.";
                label.Foreground = new SolidColorBrush(Colors.Red);
                Console.WriteLine("Error: _cachedExpiryDate is null. Ensure LicenceService.InitializeAsync is called.");
                return;
            }
            var expiryDate = _cachedExpiryDate.Value;
            var istTime = await TimeService.GetInternetTimeAsync();
            var today = istTime.Date;
            var remainingDays = (expiryDate - today).Days+1;
            if (remainingDays < 0)
            {
                remainingDays = 0;
                label.Text = $"Expired: {expiryDate:dd-MM-yyyy}";
                label.Foreground = new SolidColorBrush(Colors.Red);
            }
            else if (remainingDays == 1)
            {
                label.Text = $"Expires Today: {expiryDate:dd-MM-yyyy}";
                label.Foreground = new SolidColorBrush(Colors.Red);

            }
            else
            {
                label.Text = $"Remaining {remainingDays} Days (Expires: {expiryDate:dd-MM-yyyy})";
                label.Foreground = new SolidColorBrush(Colors.Green);
            }
        }
        catch (Exception ex)
        {
            label.Text = "Error loading expiry details.";
            label.Foreground = new SolidColorBrush(Colors.Red);
            Console.WriteLine($"Error in UpdateExpiryLabelAsync: {ex.Message}\nStackTrace: {ex.StackTrace}");
        }
    }
    #endregion

    #region Get Cached Licence Details
    public static (string EmailId, string ProductId, string MachineId, DateTime? ActivationTime) GetCachedLicenseDetails()
    {
        return (_cachedEmailId, _cachedProductId, _cachedMachineId, _cachedActivationTime);
    }
    #endregion

    #region Get Cached Info Path
    public static string GetCachedInfoPath()
    {
        return _cachedInfoPath ?? throw new InvalidOperationException("Info path not initialized. Ensure FinderService.InitializeAsync is called.");
    }
    #endregion

    #region Dispose
    public static void Dispose()
    {
        _finderCheck?.Dispose();
    }
    #endregion

    
}