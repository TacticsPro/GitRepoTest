using System.Diagnostics;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Animation;
using Microsoft.UI.Xaml.Media.Imaging;
using Office_Tools_Lite.Activation;
using Office_Tools_Lite.Contracts.Services;
using Office_Tools_Lite.Core.Contracts.Services;
using Office_Tools_Lite.Core.Services;
using Office_Tools_Lite.Models;
using Office_Tools_Lite.Services;
using Office_Tools_Lite.Task_Helper;
using Office_Tools_Lite.ViewModels;
using Office_Tools_Lite.Views;
using Windows.UI;

namespace Office_Tools_Lite;

// To learn more about WinUI 3, see https://docs.microsoft.com/windows/apps/winui/winui3/.
public partial class App : Application
{
    // The .NET Generic Host provides dependency injection, configuration, logging, and other services.
    // https://docs.microsoft.com/dotnet/core/extensions/generic-host
    // https://docs.microsoft.com/dotnet/core/extensions/dependency-injection
    // https://docs.microsoft.com/dotnet/core/extensions/configuration
    // https://docs.microsoft.com/dotnet/core/extensions/logging
    public IHost Host
    {
        get;
    }
    private static string _cachedhtmlpath;
    private DispatcherTimer _loadingTimer;
    private int _dotCount = 0;

    public static T GetService<T>()
        where T : class
    {
        if ((App.Current as App)!.Host.Services.GetService(typeof(T)) is not T service)
        {
            throw new ArgumentException($"{typeof(T)} needs to be registered in ConfigureServices within App.xaml.cs.");
        }

        return service;
    }

    public static WindowEx MainWindow { get; } = new MainWindow();

    public static UIElement? AppTitlebar
    {
        get; set;
    }

    public App()
    {
        InitializeComponent();
        Host = Microsoft.Extensions.Hosting.Host.
        CreateDefaultBuilder().
        UseContentRoot(AppContext.BaseDirectory).
        ConfigureServices((context, services) =>
        {
            // Default Activation Handler
            services.AddTransient<ActivationHandler<LaunchActivatedEventArgs>, DefaultActivationHandler>();

            // Other Activation Handlers

            // Services
            services.AddSingleton<ILocalSettingsService, LocalSettingsService>();
            services.AddSingleton<IThemeSelectorService, ThemeSelectorService>();
            services.AddTransient<INavigationViewService, NavigationViewService>();

            services.AddSingleton<IActivationService, ActivationService>();
            services.AddSingleton<IPageService, PageService>();
            services.AddSingleton<INavigationService, NavigationService>();

            // Core Services
            services.AddSingleton<IFileService, FileService>();

            // Views and ViewModels
            services.AddTransient<GSTR1_AdjustmentsViewModel>();
            services.AddTransient<GSTR1_AdjustmentsPage>();
            services.AddTransient<Tally_XMLViewModel>();
            services.AddTransient<Tally_XMLPage>();
            services.AddTransient<SettingsViewModel>();
            services.AddTransient<SettingsPage>();
            services.AddTransient<Data_EntriesViewModel>();
            services.AddTransient<Data_EntriesPage>();
            services.AddTransient<Other_ToolsViewModel>();
            services.AddTransient<Other_ToolsPage>();
            services.AddTransient<Sales_AdjustmentsViewModel>();
            services.AddTransient<Sales_AdjustmentsPage>();
            services.AddTransient<CN_DN_AdjustmentsViewModel>();
            services.AddTransient<CN_DN_AdjustmentsPage>();
            services.AddTransient<Purchase_AdjustmentsViewModel>();
            services.AddTransient<Purchase_AdjustmentsPage>();
            services.AddTransient<HomeViewModel>();
            services.AddTransient<HomePage>();
            services.AddTransient<ShellPage>();
            services.AddTransient<ShellViewModel>();

            // Configuration
            services.Configure<LocalSettingsOptions>(context.Configuration.GetSection(nameof(LocalSettingsOptions)));
        }).
        Build();

        UnhandledException += App_UnhandledException;
    }

    private void App_UnhandledException(object sender, Microsoft.UI.Xaml.UnhandledExceptionEventArgs e)
    {
        // TODO: Log and handle exceptions as appropriate.
        // https://docs.microsoft.com/windows/windows-app-sdk/api/winrt/microsoft.ui.xaml.application.unhandledexception.
    }

    protected async override void OnLaunched(LaunchActivatedEventArgs args)
    {
        base.OnLaunched(args);
        _ = RunEnableVBAAsync();
        await gettingfiles();

        // STEP 1: Show splash screen
        App.MainWindow.Content = CreateSplashContent();
        App.MainWindow.Activate();

        var cachedhtmlpath = Getting_Tutorial_Files.GettingTutorialFiles();
        _cachedhtmlpath = cachedhtmlpath;

        MainWindow.Closed += (s, e) =>
        {
            FinderService.Dispose();
            Transformation.Clear();
            _ = RunDisableVBAAsync();
        };

        // STEP 2: Wait for splash screen to be fully loaded
        var splashContent = App.MainWindow.Content as FrameworkElement;
        if (splashContent != null)
        {
            await WaitForElementLoadedAsync(splashContent);
        }

        // STEP 3: Run license service
        await FinderService.InitializeAsync(App.MainWindow);

        // STEP 4: Load full app (ShellPage)
        await App.GetService<IActivationService>().ActivateAsync(args);

    }

    private async Task WaitForElementLoadedAsync(FrameworkElement element)
    {
        if (element.IsLoaded)
        {
            return;
        }

        TaskCompletionSource<bool> tcs = new TaskCompletionSource<bool>();
        RoutedEventHandler handler = null;
        handler = (s, e) =>
        {
            element.Loaded -= handler;
            tcs.SetResult(true);
        };
        element.Loaded += handler;
        await tcs.Task;
    }

    private UIElement CreateSplashContent()
    {
        // Create the loading text block
        var loadingText = new TextBlock
        {
            Name = "LoadingText",
            Text = "Loading, please wait", // Natural phrasing, 3 dots after "Loading"
            FontFamily = new FontFamily("Cascadia Code"),
            FontSize = 24,
            FontWeight = Microsoft.UI.Text.FontWeights.Normal,
            Margin = new Thickness(0, 0, 0, 0),
            HorizontalAlignment = HorizontalAlignment.Left,
            Foreground = new SolidColorBrush(Microsoft.UI.Colors.OrangeRed)
        };

        var dottedText = new TextBlock
        {
            Name = "DottedText",
            Text = "", // Natural phrasing, 3 dots after "Loading"
            FontFamily = new FontFamily("Cascadia Code"),
            FontSize = 24,
            FontWeight = Microsoft.UI.Text.FontWeights.Normal,
            HorizontalAlignment = HorizontalAlignment.Left,
            Foreground = new SolidColorBrush(Microsoft.UI.Colors.OrangeRed),
            Opacity = 0.6
        };

        _loadingTimer = new DispatcherTimer();
        _loadingTimer.Interval = TimeSpan.FromMilliseconds(500);
        _loadingTimer.Tick += (s, e) =>
        {
            _dotCount = (_dotCount + 1) % 4; // Cycle 0–3 for standard ellipsis
            dottedText.Text = "" + new string('.', _dotCount);
        };
        _loadingTimer.Start();

        // Hook up fade animation on Loaded
        loadingText.Loaded += (s, e) =>
        {
            var fadeAnimation = new DoubleAnimation
            {
                From = 1.0,
                To = 0.4,
                Duration = new Duration(TimeSpan.FromSeconds(1.2)),
                AutoReverse = true,
                RepeatBehavior = RepeatBehavior.Forever
            };

            var storyboard = new Storyboard();
            storyboard.Children.Add(fadeAnimation);

            Storyboard.SetTarget(fadeAnimation, loadingText);
            Storyboard.SetTargetProperty(fadeAnimation, "Opacity");

            storyboard.Begin();
        };

        // Build the splash layout
        return new Grid
        {
            Children =
        {
            new Image
            {
                Source = new BitmapImage(new Uri("ms-appx:///Assets/BGImage1.jpg")),
                Stretch = Stretch.UniformToFill,
                Opacity = 0.1
            },
            new Image
            {
                Source = new BitmapImage(new Uri("ms-appx:///Assets/BGImage2.jpg")),
                Stretch = Stretch.UniformToFill,
                Opacity = 0.1
            },
            new Image
            {
                Source = new BitmapImage(new Uri("ms-appx:///Assets/BGImage3.jpg")),
                Stretch = Stretch.UniformToFill,
                Opacity = 0.1
            },
            new StackPanel
            {
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 50),
                Children =
                {
                    new TextBlock
                    {
                        Text = "Office Tools Lite",
                        FontFamily = new FontFamily("Ebrima"),
                        FontSize = 40,
                        FontWeight = Microsoft.UI.Text.FontWeights.Bold,
                        HorizontalAlignment = HorizontalAlignment.Center,
                        Foreground = new SolidColorBrush(Color.FromArgb(255, 26, 35, 126))
                    },

                    new ProgressRing
                    {
                        Width = 24,
                        Height = 24,
                        IsActive = true,
                        Margin = new Thickness(0, 0, 12, 0)
                    },

                    new StackPanel
                    {
                        Orientation = Orientation.Horizontal,
                        HorizontalAlignment = HorizontalAlignment.Center,
                        Children =
                        {
                            loadingText,
                        }
                    },
                    new Grid
                    {
                        HorizontalAlignment = HorizontalAlignment.Left,
                        Margin = new Thickness(315, -28, 0, 0),
                        Children =
                        {
                            dottedText
                        }
                    }
                }
            }

        }
        };
    }


    public static string GetCachedHtmlPath()
    {
        return _cachedhtmlpath ?? throw new InvalidOperationException("_cachedhtmlpath path not initialized. Ensure FinderService.InitializeAsync is called.");
    }

    #region get transform files
    private async Task gettingfiles()
    {
        var getfiles = new Transformation();
        await getfiles.TransformFiles();

    }
    #endregion

    #region Enable Macro
    private async Task RunEnableVBAAsync()
    {
        try
        {
            // Construct the path to the Enable_VBA_bat.bat file
            string batPath = Path.Combine(AppContext.BaseDirectory, "Runner", "Enable_VBA_bat.bat");

            // Check if the batch file exists
            if (File.Exists(batPath))
            {
                // Create a process to run the batch file
                ProcessStartInfo processInfo = new ProcessStartInfo
                {
                    FileName = batPath,
                    UseShellExecute = false,  // Do not use shell to execute (allows redirecting)
                    CreateNoWindow = true,    // Hide the console window
                    Verb = "runas"            // Runs the executable with admin privileges
                };

                // Start the process and await its exit
                using (var process = Process.Start(processInfo))
                {
                    if (process != null)
                    {
                        await process.WaitForExitAsync();
                    }
                }
            }
            else
            {
                throw new FileNotFoundException("Batch file not found: " + batPath);
            }
        }
        catch (Exception ex)
        {
            // Log or handle exceptions as needed
            Debug.WriteLine("Error running Enable_VBA_bat.bat: " + ex.Message);
        }
    }
    #endregion

    #region Disable Macro
    private async Task RunDisableVBAAsync()
    {
        try
        {
            string batPath = Path.Combine(AppContext.BaseDirectory, "Runner", "Disable_VBA_bat.bat");
            if (File.Exists(batPath))
            {
                ProcessStartInfo processInfo = new ProcessStartInfo
                {
                    FileName = batPath,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    Verb = "runas"
                };
                using (var process = Process.Start(processInfo))
                {
                    if (process != null)
                    {
                        await process.WaitForExitAsync();
                    }
                }
            }
            else
            {
                throw new FileNotFoundException("Batch file not found: " + batPath);
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine("Error running Disable_VBA_bat.bat: " + ex.Message);
        }
    }
    #endregion
}
