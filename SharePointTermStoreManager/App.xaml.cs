using System.Diagnostics;
using Microsoft.UI.Xaml;
using SharePointTermStoreManager.Services;

namespace SharePointTermStoreManager;

public partial class App : Application
{
    public static SettingsStore Settings { get; private set; } = null!;
    public static AuthService Auth { get; private set; } = null!;
    public static TermStoreService TermStore { get; private set; } = null!;
    public static MainWindow MainWindow { get; private set; } = null!;

    /// <summary>Native window handle, used to parent the MSAL sign-in.</summary>
    public static IntPtr WindowHandle { get; private set; }

    public App()
    {
        InitializeComponent();
        UnhandledException += (_, e) =>
        {
            Debug.WriteLine($"Unhandled exception: {e.Exception}");
            e.Handled = true;
        };
    }

    protected override void OnLaunched(LaunchActivatedEventArgs args)
    {
        Settings = SettingsStore.Load();
        // Removes the previous .exe an update renamed out of the way, and installed downloads.
        UpdateService.CleanUpAfterUpdate();
        // If the .exe has moved, keep its Start menu shortcut pointing at this one.
        StartMenuShortcut.KeepCurrent();
        Helpers.DisplayOptions.Current.UseColorfulIcons = Settings.Data.UseColorfulIcons;
        // Before the window exists, so the accent resources are in place from the start (see AccentColors).
        Helpers.AccentColors.Apply(Settings.Data.AccentColor);
        Auth = new AuthService();
        TermStore = new TermStoreService(Auth);

        MainWindow = new MainWindow();
        WindowHandle = WinRT.Interop.WindowNative.GetWindowHandle(MainWindow);
        MainWindow.Activate();
    }
}
