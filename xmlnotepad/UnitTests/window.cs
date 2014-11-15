using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Drawing;
using System.Reflection;
using System.Diagnostics;

namespace UnitTests {
    public class Window : IDisposable {
        Process p;
        IntPtr handle;
        bool closed;
        TestBase test;
        SystemAccessible acc;
        Dictionary<string, SystemAccessible> map;
        Window parent;
        bool disposed;

        static int delay = 100;

        public Window(Window parent, IntPtr handle) {
            this.parent = parent;
            this.handle = handle;
            this.acc = SystemAccessible.AccessibleObjectForWindow(handle);
        }

        public Window(Process p) {
            this.p = p;
            IntPtr h = p.Handle;
            while (h == IntPtr.Zero || !p.Responding) {
                Sleep(1000);
                p.WaitForInputIdle();
                h = p.Handle;
                if (p.HasExited) {
                    throw new InvalidOperationException(string.Format("Process '{0}' has exited!", p.StartInfo.FileName));
                }
            }
            p.Exited += new EventHandler(OnExited);
            int id = p.Id;
            if (handle == IntPtr.Zero) {
                // p.MainWindowHandle always returns 0 for some unknown reason...
                int retries = 20;
                while (retries-- > 0 && handle == IntPtr.Zero) {
                    handle = FindWindowForProcessId(id);
                    Sleep(1000);
                }
                if (handle == IntPtr.Zero) {
                    throw new Exception("Process as no window handle");
                }
            }
            this.acc = SystemAccessible.AccessibleObjectForWindow(handle);
        }

        void OnExited(object sender, EventArgs e) {
            if (!disposed) {
                throw new Exception("Process exited.");
            }
        }

        private static IntPtr FindWindowForProcessId(int id) {
            // Hmmm, try and find window for this process then.
            IntPtr hwnd = GetWindow(GetDesktopWindow(), GetWindowOptions.Child);
            while (hwnd != IntPtr.Zero) {
                int procid;
                int thread = GetWindowThreadProcessId(hwnd, out procid);
                if (procid == id && IsWindowVisible(hwnd)) {
                    SystemAccessible acc = SystemAccessible.AccessibleObjectForWindow(hwnd);
                    if (acc.Role == AccessibleRole.Client && 
                        (acc.State & AccessibleStates.Invisible) == 0 &&
                        !string.IsNullOrEmpty(acc.Name)) {
                        // found it!
                        return hwnd;
                    }
                }
                hwnd = GetWindow(hwnd, GetWindowOptions.Next);
            }
            return IntPtr.Zero;
        }

        public Window WaitForPopup() {
            Thread.Sleep(500);
            IntPtr h = this.Handle;
            int retries = 20;
            while (retries-- > 0) {
                IntPtr hwnd = GetForegroundWindow();
                if (hwnd != h) {
                    Trace.WriteLine("WindowChanged:" + GetForegroundWindowText());
                    Sleep(500); // give it time to get keystrokes!
                    return new Window(this, hwnd);
                }
                Sleep(1000);                
            }
            throw new ApplicationException("Window is not appearing!");
        }


        public void WaitForNewWindow() {
            int retries = 20;
            while (retries-- > 0) {
                IntPtr hwnd = GetForegroundWindow();
                if (hwnd != this.Handle) {
                    Trace.WriteLine("NewWindowOpened");
                    Sleep(1000); // give it time to get keystrokes!
                    return;
                }
                Sleep(1000);
            }
            throw new ApplicationException("Window is not appearing!");
        }

        public static IntPtr GetForegroundWindowHandle() {
            return GetForegroundWindow();
        }

        public static string GetForegroundWindowText() {
            IntPtr hwnd = GetForegroundWindow();
            if (hwnd == IntPtr.Zero)
                return "";
            return GetWindowText(hwnd);
        }

        public static string GetWindowText(IntPtr hwnd) {
            int len = GetWindowTextLength(hwnd);
            if (len <= 0) return "";
            len++; // include space for the null terminator.
            IntPtr buffer = Marshal.AllocCoTaskMem(len * 2);
            GetWindowText(hwnd, buffer, len);
            string s = Marshal.PtrToStringUni(buffer, len - 1);
            Marshal.FreeCoTaskMem(buffer);
            return s;
        }

        public void DismissPopUp(string keys) {
            Trace.WriteLine("Dismissing: " + GetForegroundWindowText());
            System.Windows.Forms.SendKeys.SendWait(keys);
            Thread.Sleep(500);
            IntPtr h = this.Handle;
            Sleep(1000);
            int retries = 10;
            while (retries-- > 0) {
                Trace.WriteLine("ForegroundWindow=" + GetForegroundWindowText());
                IntPtr hwnd = GetForegroundWindow();
                if (hwnd == h) {
                    Trace.WriteLine("WindowChanged:" + GetForegroundWindowText());
                    Sleep(500); // give it time to get keystrokes!
                    return;
                }
                Sleep(200);
            }
            throw new ApplicationException("Foreground window is not changing!");
        }

        public Window ExpectingPopup(string name) {
            Application.DoEvents();
            Window popup = this.WaitForPopup();
            string text = Window.GetForegroundWindowText();
            if (text.ToLowerInvariant() != name.ToLowerInvariant()) {
                throw new ApplicationException(string.Format("Expecting popup '{0}'", name));
            }
            return popup;
        }


        public AccessibleObject AccessibleObject {
            get {
                return this.acc;
            }
        }

        SystemAccessible GetClient() {            
            if (this.acc.Role == AccessibleRole.Client) {
                return this.acc;
            } else {
                return (SystemAccessible)this.acc.FindChildByRole(AccessibleRole.Client);
            }
        }

        public SystemAccessible FindDescendant(string name) {
            ReloadMap(name);  
            return map[name];
        }

        void ReloadMap(string name) {
            if (map == null) {
                map = new Dictionary<string, SystemAccessible>();
            }
            if (map == null || !map.ContainsKey(name)) {
                SystemAccessible client = GetClient();
                if (client != null) {
                    LoadChildren(client);
                }
            }
        }

        void LoadChildren(SystemAccessible item) {
            switch (item.Role) {
                case AccessibleRole.Alert:
                    break;
                case AccessibleRole.Animation:
                    break;
                case AccessibleRole.Application:
                    break;
                case AccessibleRole.Border:
                    break;
                case AccessibleRole.ButtonDropDown:
                    break;
                case AccessibleRole.ButtonDropDownGrid:
                    break;
                case AccessibleRole.ButtonMenu:
                    break;
                case AccessibleRole.Caret:
                    break;
                case AccessibleRole.Cell:
                    break;
                case AccessibleRole.Character:
                    break;
                case AccessibleRole.Chart:
                    break;
                case AccessibleRole.CheckButton:
                    break;
                case AccessibleRole.Client:
                case AccessibleRole.Window:
                    this.LoadComplexChildren(item);
                    break;
                case AccessibleRole.Clock:
                    break;
                case AccessibleRole.Column:
                    break;
                case AccessibleRole.ColumnHeader:
                    break;
                case AccessibleRole.ComboBox:
                    break;
                case AccessibleRole.Cursor:
                    break;
                case AccessibleRole.Default:
                    break;
                case AccessibleRole.Diagram:
                    break;
                case AccessibleRole.Dial:
                    break;
                case AccessibleRole.Dialog:
                    break;
                case AccessibleRole.Document:
                    break;
                case AccessibleRole.DropList:
                    break;
                case AccessibleRole.Equation:
                    break;
                case AccessibleRole.Graphic:
                    break;
                case AccessibleRole.Grip:
                    break;
                case AccessibleRole.Grouping:
                    break;
                case AccessibleRole.HelpBalloon:
                    break;
                case AccessibleRole.HotkeyField:
                    break;
                case AccessibleRole.Indicator:
                    break;
                case AccessibleRole.IpAddress:
                    break;
                case AccessibleRole.Link:
                    break;
                case AccessibleRole.List:
                    break;
                case AccessibleRole.ListItem:
                    break;
                case AccessibleRole.MenuBar:
                case AccessibleRole.MenuItem:
                case AccessibleRole.MenuPopup:
                case AccessibleRole.ToolBar:
                    this.LoadSimpleChildren(item);
                    break;
                case AccessibleRole.None:
                    break;
                case AccessibleRole.Outline:
                    break;
                case AccessibleRole.OutlineButton:
                    break;
                case AccessibleRole.OutlineItem:
                    break;
                case AccessibleRole.PageTab:
                    break;
                case AccessibleRole.PageTabList:
                    break;
                case AccessibleRole.Pane:
                    break;
                case AccessibleRole.ProgressBar:
                    break;
                case AccessibleRole.PropertyPage:
                    break;
                case AccessibleRole.PushButton:
                    break;
                case AccessibleRole.RadioButton:
                    break;
                case AccessibleRole.Row:
                    break;
                case AccessibleRole.RowHeader:
                    break;
                case AccessibleRole.ScrollBar:
                    break;
                case AccessibleRole.Separator:
                    break;
                case AccessibleRole.Slider:
                    break;
                case AccessibleRole.Sound:
                    break;
                case AccessibleRole.SpinButton:
                    break;
                case AccessibleRole.SplitButton:
                    break;
                case AccessibleRole.StaticText:
                    break;
                case AccessibleRole.StatusBar:
                    break;
                case AccessibleRole.Table:
                    break;
                case AccessibleRole.Text:
                    break;
                case AccessibleRole.TitleBar:
                    break;
                case AccessibleRole.ToolTip:
                    break;
                case AccessibleRole.WhiteSpace:
                    break;
                default:
                    break;
            }

        }

        void LoadComplexChildren(SystemAccessible item) {
            SystemAccessible first = item.FirstChild;
            if (first != null) {
                SystemAccessible last = item.LastChild;
                SystemAccessible child = first;
                while (child != null) {
                    AddChild(child);
                    if (child == last) break;
                    child = child.NextSibling;
                }
            }
        }

        void LoadSimpleChildren(SystemAccessible item) {
            for (int i = 1, n = item.GetChildCount(); i <= n; i++) {
                SystemAccessible child = (SystemAccessible)item.GetChild(i);
                AddChild(child);
            }
        }

        void AddChild(SystemAccessible child) {
            string name = child.Name;
            if (!string.IsNullOrEmpty(name) && child.Role != AccessibleRole.Window) {
                Trace.WriteLine("Found: " + name);
                if (!map.ContainsKey(name) || child.Role == AccessibleRole.Client) {
                    map[name] = child;
                } else {
                    Trace.WriteLine(string.Format("Ambiguous item name: '{0}'", name));
                }
            }
            if (child.GetChildCount() > 0) {
                LoadChildren(child);
            }
        }

        public IntPtr Handle {
            get {
                return handle; 
            }
        }

        public TestBase TestBase {
            get { return test; }
            set { test = value; }
        }

        public void SetWindowSize(int cx, int cy) {
            SetWindowPos(this.handle, IntPtr.Zero, 0, 0, cx, cy, (uint)SetWindowPosFlags.SWP_NOMOVE);
        }

        public void SetWindowPosition(int x, int y) {
            SetWindowPos(this.handle, IntPtr.Zero, x, y, 0, 0, (uint)SetWindowPosFlags.SWP_NOSIZE);
        }

        public void WaitForIdle(int timeout) {
            if (parent != null)
                parent.WaitForIdle(timeout);
            else if (p != null) 
                p.WaitForInputIdle();
        }

        public bool Closed {
            get { return closed; }
            set { closed = value; }
        }

        public void Sleep(int ms) {
            Thread.Sleep(ms);
        }

        public void Close() {
            if (!closed) {
                SendMessage(this.handle, (uint)WindowsMessages.WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                Sleep(delay);
            }
            closed = true;
        }

        public Window OpenDialog(string menuCommand, string formName) {
            InvokeAsyncMenuItem(menuCommand);
            return this.WaitForPopup();
        }

        public void SendKeys(string keys) {
            if (Window.GetForegroundWindow() != this.handle) {
                Activate();
                Sleep(200);
            }
            System.Windows.Forms.SendKeys.SendWait(keys);
        }

        public void Activate() {
            SetForegroundWindow(this.handle);
        }

        const uint OBJID_CLIENT = 0xFFFFFFFC;

        public void InvokeMenuItem(string menuItemName) {
            this.ReloadMap(menuItemName);
            Sleep(30);
            this.WaitForIdle(2000);

            SystemAccessible item;
            if (!this.map.TryGetValue(menuItemName, out item)) {
                throw new Exception(string.Format("Menu item '{0}' not found", menuItemName));                
            }
            Trace.WriteLine("InvokeMenuItem(" + menuItemName + ")");
            item.DoDefaultAction();
        }

        public void InvokeAsyncMenuItem(string menuItemName) {
            this.ReloadMap(menuItemName);
            Sleep(delay);
            this.WaitForIdle(2000);

            SystemAccessible item;
            if (!this.map.TryGetValue(menuItemName, out item)) {
                throw new Exception(string.Format("Menu item '{0}' not found", menuItemName));                
            }
            Trace.WriteLine("InvokeAsyncMenuItem(" + menuItemName + ")");

            // this is NOT async with things like SaveAs where a dialog pops up!
            //item.DoDefaultAction();

            this.WaitForIdle(1000);
            Sleep(1000);
            if (item.Role == AccessibleRole.PushButton) {
                // Toolbar buttons cannot be invoked from shortcuts.
                Mouse.MouseClick(Center(item.Bounds), MouseButtons.Left);
            } else {
                TypeShortcut(item);
            }
        }

        Point Center(Rectangle bounds) {
            return new Point(bounds.Left + (bounds.Width / 2),
                bounds.Top + (bounds.Height / 2));
        }


        void TypeShortcut(AccessibleObject item) {
            Sleep(1000);
            AccessibleObject parent = item.Parent;
            if (parent.Role == AccessibleRole.MenuItem) {
                TypeShortcut(parent);
            }
            string shortcut = item.KeyboardShortcut;
            if (!string.IsNullOrEmpty(shortcut)) {
                SendShortcut(shortcut);
            } else {
                throw new NotImplementedException("InvokeAsyncMenuItem can't work without menu item shortcuts");
            }
        }

        public void SendShortcut(string shortcut) {
            string keys = shortcut.Replace("Alt+", "%").Replace("Ctrl+", "^").Replace("Shift+", "+");
            Sleep(200);
            SendKeys(keys);
        }

        public Rectangle GetScreenBounds() {
            WINDOWINFO wi = new WINDOWINFO();
            wi.cbSize = Marshal.SizeOf(wi);
            if (GetWindowInfo(this.handle, ref wi)) {
                RECT r = wi.rcWindow;
                return new Rectangle(r.left, r.top, r.right - r.left, r.bottom - r.top);
            }
            return Rectangle.Empty;
        }

        #region IDisposable Members

        public void Dispose() {
            disposed = true;
            if (p != null && !p.HasExited) {
                this.Close();
                Sleep(1000);
                if (!p.HasExited) {
                    p.Kill();
                }                
            }
        }

        #endregion

        enum GetWindowOptions {
            First    = 0,
            Last = 1,
            Next = 2,
            Previous = 3,
            Owner = 5,
            Child = 5
        }

        [DllImport("User32")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("User32")]
        public static extern IntPtr GetTopWindow(IntPtr hwnd);

        [DllImport("User32")]
        public static extern IntPtr GetDesktopWindow();

        [DllImport("User32")]
        public static extern IntPtr GetLastActivePopup(IntPtr hWnd);

        [DllImport("User32", CharSet = CharSet.Unicode)]
        static extern int GetWindowTextLength(IntPtr hwnd);

        [DllImport("User32", CharSet = CharSet.Unicode)]
        static extern int GetWindowText(IntPtr hWnd, IntPtr lpString, int nMaxCount);

        [DllImport("User32", CharSet = CharSet.Unicode)]
        static extern IntPtr GetWindow(IntPtr hWnd, GetWindowOptions uCmd);

        [DllImport("User32", CharSet = CharSet.Unicode)]
        static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        [DllImport("User32", CharSet = CharSet.Unicode)]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("User32", CharSet = CharSet.Unicode)]
        static extern bool GetWindowInfo(IntPtr hwnd, ref WINDOWINFO pwi);

        [DllImport("User32", CharSet = CharSet.Unicode)]
        static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("User32", CharSet = CharSet.Unicode)]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShows);

        [DllImport("User32", CharSet = CharSet.Unicode)]
        static extern void SetForegroundWindow(IntPtr hWnd);

        enum WindowsMessages {
            WM_CLOSE = 0x0010
        }

        [DllImport("User32", CharSet = CharSet.Unicode)]
        static extern int SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        struct WINDOWINFO {
            public int cbSize;
            public RECT rcWindow;
            public RECT rcClient;
            public int dwStyle;
            public int dwExStyle;
            public int dwWindowStatus;
            public uint cxWindowBorders;
            public uint cyWindowBorders;
            public IntPtr atomWindowType;
            public int wCreatorVersion;
        }

        enum SetWindowPosFlags {
            SWP_NOSIZE = 0x0001,
            SWP_NOMOVE = 0x0002,
            SWP_NOZORDER = 0x0004,
            SWP_NOREDRAW = 0x0008,
            SWP_NOACTIVATE = 0x0010,
            SWP_FRAMECHANGED = 0x0020,  /* The frame changed: send WM_NCCALCSIZE */
            SWP_SHOWWINDOW = 0x0040,
            SWP_HIDEWINDOW = 0x0080,
            SWP_NOCOPYBITS = 0x0100,
            SWP_NOOWNERZORDER = 0x0200, /* Don't do owner Z ordering */
            SWP_NOSENDCHANGING = 0x0400,  /* Don't send WM_WINDOWPOSCHANGING */
            SWP_DRAWFRAME = SWP_FRAMECHANGED,
            SWP_NOREPOSITION = SWP_NOOWNERZORDER
        }

        enum ShowWindowFlags {
            SW_HIDE = 0,
            SW_SHOWNORMAL = 1,
            SW_NORMAL = 1,
            SW_SHOWMINIMIZED = 2,
            SW_SHOWMAXIMIZED = 3,
            SW_MAXIMIZE = 3,
            SW_SHOWNOACTIVATE = 4,
            SW_SHOW = 5,
            SW_MINIMIZE = 6,
            SW_SHOWMINNOACTIVE = 7,
            SW_SHOWNA = 8,
            SW_RESTORE = 9,
            SW_SHOWDEFAULT = 10,
            SW_FORCEMINIMIZE = 11,
            SW_MAX = 11
        }
    }
}
