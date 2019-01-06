using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace BbgTerminalLogin
{

    public class Program
    {
        private const string _bbgWindowCaptionRegex = @"^\d+-BLOOMBERG$";

        private static TimeSpan Delay = TimeSpan.FromSeconds(2);

        private static int Main(string[] args)
        {
            string username = "USERNAME"; // <-- BBG USER NAME GOES HERE
            string password = "PASSWORD"; // <-- BBG PASSWORD GOES HERE

            foreach (var window in WindowInfo.EnumWindows())
            {
                if (Regex.IsMatch(window.WindowCaption, _bbgWindowCaptionRegex, RegexOptions.IgnoreCase))
                {
                    Log.LogInformation(() => "Evaluating [" + window.WindowCaption + "].");
                    if (!window.WindowStyles.HasFlag(WindowStyles.VISIBLE))
                    {
                        Log.LogWarning(() => "Skipping [" + window.WindowCaption + "] because it's not visible.");
                        continue;
                    }
                    var blpCtrlWin2Count = GetBLPControlWindow2Count(window);
                    if (blpCtrlWin2Count >= 2)
                    {
                        Log.LogInformation(() => "Already logged in from [" + window.WindowCaption + "]. The window count is {blpCtrlWin2Count}.");
                        return (int)ExitCodes.AlreadyLoggedIn;
                    }
                }
            }
            foreach (var window in WindowInfo.EnumWindows().OrderBy(x => x.WindowCaption))
            {
                if (window.ClassName == "Notepad")
                {
                    // for debugging locally w/ notepad.
                    PerformLogin(window, username, password); return (int) ExitCodes.LoginPerformed;
                }
                if (Regex.IsMatch(window.WindowCaption, _bbgWindowCaptionRegex, RegexOptions.IgnoreCase))
                {
                    Log.LogInformation(() => "Evaluating [" + window.WindowCaption + "].");
                    if (!window.WindowStyles.HasFlag(WindowStyles.VISIBLE))
                    {
                        Log.LogWarning(() => "Skipping [" + window.WindowCaption + "] because it's not visible.");
                        continue;
                    }
                    var blpCtrlWin2Count = GetBLPControlWindow2Count(window);
                    if (blpCtrlWin2Count == 1)
                    {
                        PerformLogin(window, username, password);
                        return (int)ExitCodes.LoginPerformed;
                    }
                }
            }
            throw new Exception("Bloomberg window was not found.");
        }

        private static void PerformLogin(WindowInfo window, string userName, string password)
        {
            Log.LogInformation(() => "logging in to [" + window.WindowCaption + "].");
            var bloombergWindow = IntPtr.Zero;

            for (int i = 0; i < 2; i++)
            {
                //we do ESC+ESC+ENTER to go to the logon screen
                SendInput.PressKeys(window.Handle, ScanCodes.Escape, ScanCodes.Escape, ScanCodes.Enter);

                //bloomberg may switch to a different window to do the logon so we capture that window.
                bloombergWindow = NativeMethods.GetForegroundWindow();
                Thread.Sleep(Delay);

                //we do it twice to ensure we got the right window for username/password input.
            }
            if (bloombergWindow == IntPtr.Zero)
            {
                throw new Exception("A window was not activated.");
            }
            bloombergWindow = window.Handle;

            var activeBbgWindow = FindTopLevelOwner(bloombergWindow);
            if (activeBbgWindow == null)
            {
                throw new Exception("The window for the handle [" + bloombergWindow + "] could not be found.");
            }
            Log.LogInformation(() => "The [" + activeBbgWindow.WindowCaption + "] window was activated.");

            SendInput.PressKeys(activeBbgWindow.Handle, userName);
            SendInput.PressKeys(activeBbgWindow.Handle, ScanCodes.Tab);

            Thread.Sleep(Delay);

            SendInput.PressKeys(activeBbgWindow.Handle, password);
            SendInput.PressKeys(activeBbgWindow.Handle, ScanCodes.Enter);
        }

        private static WindowInfo FindTopLevelOwner(IntPtr handle)
        {
            return WindowInfo.EnumWindows()
                .Where(w => (w.Handle == handle) || w.GetChildWindows().Any(c => c.Handle == handle))
                .SingleOrDefault();
        }

        private static int GetBLPControlWindow2Count(WindowInfo window)
        {
            return window.GetChildWindows()
                .Where(w => w.ClassName == "BLPControlWindow2")
                .Where(w => w.WindowStyles.HasFlag(WindowStyles.VISIBLE))
                .Where(w => w.Parent.WindowStyles.HasFlag(WindowStyles.VISIBLE))
                .Count();
        }
    }

    [Flags()]
    public enum WindowStyles : uint
    {
        /// <summary>The window has a thin-line border.</summary>
        BORDER = 0x800000,

        /// <summary>The window has a title bar (includes the BORDER style).</summary>
        CAPTION = 0xc00000,

        /// <summary>The window is a child window. A window with this style cannot have a menu bar. This style cannot be used with the POPUP style.</summary>
        CHILD = 0x40000000,

        /// <summary>Excludes the area occupied by child windows when drawing occurs within the parent window. This style is used when creating the parent window.</summary>
        CLIPCHILDREN = 0x2000000,

        /// <summary>
        /// Clips child windows relative to each other; that is, when a particular child window receives a WM_PAINT message, the CLIPSIBLINGS style clips all other overlapping child windows out of the region of the child window to be updated.
        /// If CLIPSIBLINGS is not specified and child windows overlap, it is possible, when drawing within the client area of a child window, to draw within the client area of a neighboring child window.
        /// </summary>
        CLIPSIBLINGS = 0x4000000,

        /// <summary>The window is initially disabled. A disabled window cannot receive input from the user. To change this after a window has been created, use the EnableWindow function.</summary>
        DISABLED = 0x8000000,

        /// <summary>The window has a border of a style typically used with dialog boxes. A window with this style cannot have a title bar.</summary>
        DLGFRAME = 0x400000,

        /// <summary>
        /// The window is the first control of a group of controls. The group consists of this first control and all controls defined after it, up to the next control with the GROUP style.
        /// The first control in each group usually has the TABSTOP style so that the user can move from group to group. The user can subsequently change the keyboard focus from one control in the group to the next control in the group by using the direction keys.
        /// You can turn this style on and off to change dialog box navigation. To change this style after a window has been created, use the SetWindowLong function.
        /// </summary>
        GROUP = 0x20000,

        /// <summary>The window has a horizontal scroll bar.</summary>
        HSCROLL = 0x100000,

        /// <summary>The window is initially maximized.</summary> 
        MAXIMIZE = 0x1000000,

        /// <summary>The window has a maximize button. Cannot be combined with the EX_CONTEXTHELP style. The SYSMENU style must also be specified.</summary> 
        MAXIMIZEBOX = 0x10000,

        /// <summary>The window is initially minimized.</summary>
        MINIMIZE = 0x20000000,

        /// <summary>The window has a minimize button. Cannot be combined with the EX_CONTEXTHELP style. The SYSMENU style must also be specified.</summary>
        MINIMIZEBOX = 0x20000,

        /// <summary>The window is an overlapped window. An overlapped window has a title bar and a border.</summary>
        OVERLAPPED = 0x0,

        /// <summary>The window is an overlapped window.</summary>
        OVERLAPPEDWINDOW = OVERLAPPED | CAPTION | SYSMENU | SIZEFRAME | MINIMIZEBOX | MAXIMIZEBOX,

        /// <summary>The window is a pop-up window. This style cannot be used with the CHILD style.</summary>
        POPUP = 0x80000000u,

        /// <summary>The window is a pop-up window. The CAPTION and POPUPWINDOW styles must be combined to make the window menu visible.</summary>
        POPUPWINDOW = POPUP | BORDER | SYSMENU,

        /// <summary>The window has a sizing border.</summary>
        SIZEFRAME = 0x40000,

        /// <summary>The window has a window menu on its title bar. The CAPTION style must also be specified.</summary>
        SYSMENU = 0x80000,

        /// <summary>
        /// The window is a control that can receive the keyboard focus when the user presses the TAB key.
        /// Pressing the TAB key changes the keyboard focus to the next control with the TABSTOP style.  
        /// You can turn this style on and off to change dialog box navigation. To change this style after a window has been created, use the SetWindowLong function.
        /// For user-created windows and modeless dialogs to work with tab stops, alter the message loop to call the IsDialogMessage function.
        /// </summary>
        TABSTOP = 0x10000,

        /// <summary>The window is initially visible. This style can be turned on and off by using the ShowWindow or SetWindowPos function.</summary>
        VISIBLE = 0x10000000,

        /// <summary>The window has a vertical scroll bar.</summary>
        VSCROLL = 0x200000,
    }

    [return: MarshalAs(UnmanagedType.Bool)]
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [return: MarshalAs(UnmanagedType.Bool)]
    public delegate bool EnumChildProc(IntPtr hWnd, IntPtr lParam);

    public enum GWL
    {
        EXSTYLE = -20,
        HINSTANCE = -6,
        HWNDPARENT = -8,
        ID = -12,
        STYLE = -16,
        USERDATA = -21,
        WNDPROC = -4,
    }

    public enum ScanCodes : ushort
    {
        Enter = 0x1C,
        Tab = 0x0F,
        Delete = 0x0E,
        Escape = 0x01,
    }

    public enum InputEventType : uint
    {
        Mouse = 0,
        Keyboard = 1,
        Hardware = 2,
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct HARDWAREINPUT
    {
        public int uMsg;

        public short wParamL;

        public short wParamH;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct MOUSEINPUT
    {
        public int dx;

        public int dy;

        public int mouseData;

        public int dwFlags;

        public int time;

        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Explicit)]
    public struct INPUTUNION
    {
        [FieldOffset(0)]
        public MOUSEINPUT mi;

        [FieldOffset(0)]
        public KEYBDINPUT ki;

        [FieldOffset(0)]
        public HARDWAREINPUT hi;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct INPUT
    {
        public InputEventType type;

        public INPUTUNION inputUnion;
    }

    [Flags]
    public enum KEYEVENTF : uint
    {
        NONE = 0x0,
        EXTENDEDKEY = 0x1,
        KEYUP = 0x2,
        UNICODE = 0x4,
        SCANCODE = 0x8,
    }

    public enum VirtualKeys : ushort
    {
        Enter = 13,
        Tab = 9,
        Escape = 27,
        Home = 36,
        End = 35,
        Left = 37,
        Right = 39,
        Up = 38,
        Down = 40,
        PageUp = 33,
        PageDown = 34,
        NumLock = 144,
        ScrollLock = 145,
        PrintScreen = 44,
        Break = 3,
        Backspace = 8,
        CLear = 12,
        CapsLock = 20,
        Insert = 45,
        Delete = 46,
        Help = 47,
        F1 = 112,
        F2 = 113,
        F3 = 114,
        F4 = 115,
        F5 = 116,
        F6 = 117,
        F7 = 118,
        F8 = 119,
        F9 = 120,
        F10 = 121,
        F11 = 122,
        F12 = 123,
        F13 = 124,
        F14 = 125,
        F15 = 126,
        F16 = 127,
        Multiply = 106,
        Add = 107,
        Subtract = 109,
        Divide = 111,
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct KEYBDINPUT
    {
        public VirtualKeys wVk;

        public ushort wScan;

        public KEYEVENTF dwFlags;

        public int time;

        public IntPtr dwExtraInfo;
    }

    public static class SendInput
    {
        public static void PressKey(IntPtr? window, INPUT input)
        {
            if (window.HasValue)
            {
                NativeMethods.SetForegroundWindow(window.Value);
            }
            var result = NativeMethods.SendInput(1, ref input, Marshal.SizeOf(typeof(INPUT)));
            if (result == 0)
            {
                throw new Win32Exception();
            }
            Thread.Sleep(1);

            input.inputUnion.ki.dwFlags |= KEYEVENTF.KEYUP;
            if (window.HasValue)
            {
                NativeMethods.SetForegroundWindow(window.Value);
            }
            result = NativeMethods.SendInput(1, ref input, Marshal.SizeOf(typeof(INPUT)));
            if (result == 0)
            {
                throw new Win32Exception();
            }
            Thread.Sleep(1);
        }

        public static void PressKeys(IntPtr? window, string keys)
        {
            foreach (var key in keys)
            {
                var input = new INPUT
                {
                    type = InputEventType.Keyboard,
                    inputUnion = new INPUTUNION
                    {
                        ki = new KEYBDINPUT
                        {
                            wScan = key,
                            dwFlags = KEYEVENTF.UNICODE,
                        }
                    }
                };
                PressKey(window, input);
            }
        }

        public static void PressKeys(IntPtr? window, params VirtualKeys[] keys)
        {
            foreach (var key in keys)
            {
                var input = new INPUT
                {
                    type = InputEventType.Keyboard,
                    inputUnion = new INPUTUNION
                    {
                        ki = new KEYBDINPUT
                        {
                            wVk = key,
                            dwFlags = KEYEVENTF.NONE,
                        }
                    }
                };
                PressKey(window, input);
            }
        }

        public static void PressKeys(IntPtr? window, params ScanCodes[] codes)
        {
            foreach (var code in codes)
            {
                var input = new INPUT
                {
                    type = InputEventType.Keyboard,
                    inputUnion = new INPUTUNION
                    {
                        ki = new KEYBDINPUT
                        {
                            wScan = (ushort)code,
                            dwFlags = KEYEVENTF.SCANCODE,
                        }
                    }
                };
                PressKey(window, input);
            }
        }
    }

    public static class NativeMethods
    {
        [DllImport("user32.dll", SetLastError = true)]
        public static extern int GetWindowLong(IntPtr hwnd, GWL nIndex);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        public static extern IntPtr GetParent(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumChildWindows(IntPtr hwnd, EnumChildProc func, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern int GetWindowText(IntPtr hwnd, StringBuilder lpString, long cch);

        [DllImport("user32.dll")]
        public static extern IntPtr GetMessageExtraInfo();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint SendInput(uint nInputs, ref INPUT pInput, int cbSize);

        [DllImport("user32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr lpEnumFunc, out int lpdwProcessId);
    }

    [DebuggerDisplay("{Handle.ToInt64().ToString(\"x\")} - {WindowCaption}")]
    public class WindowInfo
    {
        private int _procId;

        public WindowInfo(IntPtr hwnd)
        {
            Handle = hwnd;
            NativeMethods.GetWindowThreadProcessId(hwnd, out _procId);

            var buffer = new StringBuilder(65536);
            NativeMethods.GetClassName(hwnd, buffer, buffer.Capacity);
            ClassName = buffer.ToString();
        }

        public IntPtr Handle;

        public int ProcessId { get { return _procId; } }

        public WindowStyles WindowStyles
        {
            get
            {
                var style = NativeMethods.GetWindowLong(Handle, GWL.STYLE);
                if (style == 0)
                {
                    throw new Win32Exception();
                }
                return (WindowStyles)style;
            }
        }

        public WindowInfo Parent
        {
            get
            {
                var parent = NativeMethods.GetParent(Handle);
                if (parent == IntPtr.Zero)
                {
                    throw new Win32Exception();
                }
                return new WindowInfo(parent);
            }
        }

        public string WindowCaption
        {
            get
            {
                var length = NativeMethods.GetWindowTextLength(Handle);
                if (length == 0)
                {
                    return string.Empty;
                }
                var buffer = new StringBuilder(length + 1);
                NativeMethods.GetWindowText(Handle, buffer, buffer.Capacity);
                return buffer.ToString();
            }
        }

        public string ClassName;

        public static IEnumerable<WindowInfo> EnumWindows()
        {
            var windows = new List<WindowInfo>();
            EnumWindowsProc callback = (hwnd, lparam) =>
            {
                windows.Add(new WindowInfo(hwnd));
                return true;
            };
            if (!NativeMethods.EnumWindows(callback, IntPtr.Zero))
            {
                throw new Win32Exception();
            }
            return windows;
        }

        public IEnumerable<WindowInfo> GetChildWindows()
        {
            var windows = new List<WindowInfo>();
            EnumChildProc callback = (hwnd, lparam) =>
            {
                windows.Add(new WindowInfo(hwnd));
                return true;
            };
            NativeMethods.EnumChildWindows(Handle, callback, IntPtr.Zero);
            return windows;
        }

        public void SetForegroundWindow()
        {
            if (!NativeMethods.SetForegroundWindow(Handle))
            {
                throw new Win32Exception("Unable to set to foreground.");
            }
        }
    }

    public enum ExitCodes
    {
        LoginPerformed = 0,
        AlreadyLoggedIn = 16000,
    }

    public static class Log
    {
        public static void LogInformation(Func<string> message)
        {
            Console.WriteLine(string.Format("INFO: {0}: {1}", DateTime.Now, message()));
        }

        public static void LogWarning(Func<string> message)
        {
            Console.WriteLine(string.Format("WARN: {0}: {1}", DateTime.Now, message()));
        }
    }
}
