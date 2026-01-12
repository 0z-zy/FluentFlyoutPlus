using FluentFlyout.Classes.Settings;
using System.Runtime.InteropServices;
using System.Text;

namespace FluentFlyoutWPF.Classes;

internal class FullscreenDetector
{
    private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();

    // Window placement constants
    private const int SW_SHOWMAXIMIZED = 3;

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left, Top, Right, Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT
    {
        public int X, Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct WINDOWPLACEMENT
    {
        public int length;
        public int flags;
        public int showCmd;
        public POINT ptMinPosition;
        public POINT ptMaxPosition;
        public RECT rcNormalPosition;
        public RECT rcDevice;
    }

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    private static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    private const uint MONITOR_DEFAULTTONEAREST = 2;
    
    /// <summary>
    /// Represents the different states of user notification returned by the Windows Shell API.
    /// </summary>
    public enum QUERY_USER_NOTIFICATION_STATE
    {
        QUNS_NOT_PRESENT = 1,
        QUNS_BUSY = 2,
        QUNS_RUNNING_D3D_FULL_SCREEN = 3,
        QUNS_PRESENTATION_MODE = 4,
        QUNS_ACCEPTS_NOTIFICATIONS = 5,
        QUNS_QUIET_TIME = 6,
        QUNS_APP = 7
    }

    [DllImport("shell32.dll")]
    private static extern int SHQueryUserNotificationState(out QUERY_USER_NOTIFICATION_STATE pquns);

    /// <summary>
    /// Checks if a DirectX exclusive fullscreen application or game is currently running.
    /// </summary>
    /// <returns>
    /// true if a fullscreen DirectX application is running;
    /// false if no fullscreen application is detected, DisableIfFullscreen setting is false, or if the check fails
    /// </returns>
    public static bool IsFullscreenApplicationRunning()
    {
        if (!SettingsManager.Current.DisableIfFullscreen) return false;
        try
        {
            QUERY_USER_NOTIFICATION_STATE state;
            int result = SHQueryUserNotificationState(out state);

            if (result != 0) // 0 means SUCCESS
            {
                throw new Exception($"SHQueryUserNotificationState failed with error code: {result}");
            }

            return state == QUERY_USER_NOTIFICATION_STATE.QUNS_RUNNING_D3D_FULL_SCREEN;
        }
        catch (Exception ex)
        {
            Logger.Error(ex, "Error detecting fullscreen state");
            return false;
        }
    }

    /// <summary>
    /// Checks if the foreground window is maximized or in fullscreen mode (non-DirectX).
    /// This is used to hide the taskbar widget when apps cover the taskbar area.
    /// </summary>
    /// <param name="targetMonitor">Optional: The monitor handle to check against. If provided, only returns true if the foreground window is on this monitor.</param>
    /// <returns>true if the foreground window is maximized or fullscreen; false otherwise</returns>
    public static bool IsForegroundWindowMaximizedOrFullscreen(IntPtr targetMonitor = default)
    {
        try
        {
            IntPtr foregroundWindow = GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero)
                return false;

            // Ignore shell windows (taskbar, desktop, etc.)
            StringBuilder className = new(256);
            GetClassName(foregroundWindow, className, className.Capacity);
            string windowClass = className.ToString();

            // List of shell window classes to ignore
            if (windowClass == "Shell_TrayWnd" ||           // Main taskbar
                windowClass == "Shell_SecondaryTrayWnd" ||  // Secondary monitor taskbar
                windowClass == "Progman" ||                 // Desktop
                windowClass == "WorkerW" ||                 // Desktop worker
                windowClass == "Windows.UI.Core.CoreWindow" || // Start menu, action center
                windowClass == "XamlExplorerHostIslandWindow") // Modern shell elements
            {
                return false;
            }

            // Check if the foreground window is on the target monitor (if specified)
            if (targetMonitor != IntPtr.Zero)
            {
                IntPtr windowMonitor = MonitorFromWindow(foregroundWindow, MONITOR_DEFAULTTONEAREST);
                if (windowMonitor != targetMonitor)
                    return false;
            }

            // Check if window is maximized
            WINDOWPLACEMENT placement = new() { length = Marshal.SizeOf<WINDOWPLACEMENT>() };
            if (GetWindowPlacement(foregroundWindow, ref placement))
            {
                if (placement.showCmd == SW_SHOWMAXIMIZED)
                    return true;
            }

            // Also check notification state for presentation mode and D3D fullscreen
            if (SHQueryUserNotificationState(out QUERY_USER_NOTIFICATION_STATE state) == 0)
            {
                if (state == QUERY_USER_NOTIFICATION_STATE.QUNS_RUNNING_D3D_FULL_SCREEN ||
                    state == QUERY_USER_NOTIFICATION_STATE.QUNS_PRESENTATION_MODE)
                {
                    return true;
                }
            }

            return false;
        }
        catch (Exception ex)
        {
            Logger.Error(ex, "Error detecting maximized/fullscreen state");
            return false;
        }
    }
}