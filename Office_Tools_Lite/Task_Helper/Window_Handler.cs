using System.Runtime.InteropServices;
using Microsoft.UI.Xaml;
using WinRT.Interop;

namespace Office_Tools_Lite.Task_Helper;
public static class Window_Handler
{
    [DllImport("User32.dll")]
    public static extern bool ShowWindow(nint hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(nint hWnd);

    public const int SW_MINIMIZE = 6;
    public const int SW_RESTORE = 9;

    public static void Minimize(Window mainwindow)
    {
        var hWnd = WindowNative.GetWindowHandle(mainwindow);
        ShowWindow(hWnd, SW_MINIMIZE);
    }

    public static void Restore(Window mainwindow)
    {
        var hWnd = WindowNative.GetWindowHandle(mainwindow);
        ShowWindow(hWnd, SW_RESTORE);
    }

    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    
}
