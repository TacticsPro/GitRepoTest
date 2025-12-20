using CommunityToolkit.Mvvm.ComponentModel;

using Microsoft.UI.Xaml.Controls;

using Office_Tools_Lite.Contracts.Services;
using Office_Tools_Lite.ViewModels;
using Office_Tools_Lite.Views;

namespace Office_Tools_Lite.Services;

public class PageService : IPageService
{
    private readonly Dictionary<string, Type> _pages = new();

    public PageService()
    {
        Configure<HomeViewModel, HomePage>();
        Configure<Purchase_AdjustmentsViewModel, Purchase_AdjustmentsPage>();
        Configure<CN_DN_AdjustmentsViewModel, CN_DN_AdjustmentsPage>();
        Configure<Sales_AdjustmentsViewModel, Sales_AdjustmentsPage>();
        Configure<Other_ToolsViewModel, Other_ToolsPage>();
        Configure<Data_EntriesViewModel, Data_EntriesPage>();
        Configure<SettingsViewModel, SettingsPage>();
        Configure<Tally_XMLViewModel, Tally_XMLPage>();
        Configure<GSTR1_AdjustmentsViewModel, GSTR1_AdjustmentsPage>();
    }

    public Type GetPageType(string key)
    {
        Type? pageType;
        lock (_pages)
        {
            if (!_pages.TryGetValue(key, out pageType))
            {
                throw new ArgumentException($"Page not found: {key}. Did you forget to call PageService.Configure?");
            }
        }

        return pageType;
    }

    private void Configure<VM, V>()
        where VM : ObservableObject
        where V : Page
    {
        lock (_pages)
        {
            var key = typeof(VM).FullName!;
            if (_pages.ContainsKey(key))
            {
                throw new ArgumentException($"The key {key} is already configured in PageService");
            }

            var type = typeof(V);
            if (_pages.ContainsValue(type))
            {
                throw new ArgumentException($"This type is already configured with key {_pages.First(p => p.Value == type).Key}");
            }

            _pages.Add(key, type);
        }
    }
}
