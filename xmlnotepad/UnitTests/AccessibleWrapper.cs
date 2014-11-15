using System;
using System.Collections.Generic;
using System.Text;
using Accessibility;
using System.Windows.Forms;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.IO;
using System.Diagnostics;

namespace UnitTests {

    public class FileComparer : IComparer<string> {
        public int Compare(string a, string b) {
           if (a == null && b == null) return 0;
            if (a == null) return -1;
            if (b == null) return 1;
            // The operating system may or may not have visible "file extensions" so we do an extensionless compare
            return string.Compare(Path.GetFileNameWithoutExtension(a), Path.GetFileNameWithoutExtension(b), StringComparison.CurrentCultureIgnoreCase);
        }
    }

    public class StringComparer : IComparer<string> {
        public int Compare(string a, string b) {
            return String.Compare(a, b, StringComparison.CurrentCultureIgnoreCase);
        }
    }

    public class SystemAccessible : AccessibleObject {
        IAccessible acc;
        int id;
        AccessibleObject parent;
        AccessibleObject firstChild;
        AccessibleObject lastChild;
        AccessibleObject next;
        static Dictionary<object, SystemAccessible> s_Wrappers = new Dictionary<object, SystemAccessible>();

        public SystemAccessible(AccessibleObject parent, IAccessible acc, int id) {
            this.acc = acc;
            this.id = id;
            this.parent = parent;
        }

        public override string Name {
            get {
                try {
                    if (this.Role == AccessibleRole.WhiteSpace) return null;
                    return acc.get_accName(id);
                } catch {
                    return null;
                }
            }
            set {
                acc.set_accName(id, value);
            }
        }

        public override AccessibleRole Role {
            get {
                try {
                    int role = (int)acc.get_accRole(id);
                    return (AccessibleRole)role;
                } catch {
                }
                return AccessibleRole.WhiteSpace;
            }
        }

        public override AccessibleObject GetFocused() {
            try {
                object focus = acc.accFocus;                
                IAccessible a = focus as IAccessible;
                if (a != null) return Wrap(a);
                if (focus is int) {
                    return new SystemAccessible(this, this.acc, (int)focus);
                }
            } catch {
            }
            return null;
        }

        public override string DefaultAction {
            get {
                try {
                    return acc.get_accDefaultAction(id);
                } catch {
                    return null;
                }
            }
        }

        public override void DoDefaultAction() {
            try {
                acc.accDoDefaultAction(id);
            } catch {
            }
        }

        public override string Description {
            get {
                try {
                    return acc.get_accDescription(id);
                } catch {
                    return null;
                }
            }
        }

        public override string Help {
            get {
                try {
                    return acc.get_accHelp(id);
                } catch {
                    return null;
                }
            }
        }

        public override AccessibleStates State {
            get {
                try {
                    int i = (int)acc.get_accState(id);
                    return (AccessibleStates)i;
                } catch {
                    return AccessibleStates.None;
                }
            }
        }

        public override string Value {
            get {
                try {
                    return acc.get_accValue(id);
                } catch {
                    return null;
                }
            }
            set {
                acc.set_accValue(id, value);
            }
        }

        public IntPtr Hwnd {
            get {
                IntPtr hwnd;
                WindowFromAccessibleObject(this.acc, out hwnd);
                return hwnd;
            }
        }

        public override Rectangle Bounds {
            get {
                try {
                    int left, top, width, height;
                    acc.accLocation(out left, out top, out width, out height, id);
                    return new Rectangle(left, top, width, height);
                } catch {
                }
                return Rectangle.Empty;
            }
        }

        public override int GetChildCount() {
            if (id == CHILDID_SELF) {
                try {
                    if (this.Role == AccessibleRole.WhiteSpace) return 0;
                    return acc.accChildCount;
                } catch {
                    // ignore errors.
                }
            }
            return 0;
        }

        public override AccessibleObject GetChild(int index) {
            try {
                object o = acc.get_accChild(index);
                if (o == null && this.GetChildCount() > 0) {
                    return new SystemAccessible(this, this.acc, index);
                }
                return Wrap(o);
            } catch {
                return null;
            }
        }

        AccessibleObject NavigateToChild(AccessibleObject acc, int index) {
            AccessibleObject c = acc.Navigate(AccessibleNavigation.FirstChild);
            while (--index > 0 && c != null) {
                c = c.Navigate(AccessibleNavigation.Next);
            }
            return c;
        }

        public override AccessibleObject GetSelected() {
            try {
                object s = acc.accSelection;
                return Wrap(s);
            } catch {
                return null;
            }
        }

        public override int GetHelpTopic(out string fileName) {
            try {
                return acc.get_accHelpTopic(out fileName, this.id);                
            } catch {
                fileName = null;
                return -1;
            }
        }

        public override AccessibleObject HitTest(int x, int y) {
            try {
                object o = acc.accHitTest(x, y);
                return Wrap(o);
            } catch (Exception) {
                return null;
            }            
        }

        public override string KeyboardShortcut {
            get {
                try {
                    return acc.get_accKeyboardShortcut(this.id);
                } catch (Exception) {
                    return null;
                }
            }
        }

        public override AccessibleObject Navigate(AccessibleNavigation navdir) {
            try {
                object o = acc.accNavigate((int)navdir, this.id);
                return Wrap(o);
            } catch (Exception) {
                return null;
            }
        }

        public override void Select(AccessibleSelection flags) {
            acc.accSelect((int)flags, this.id);
        }

        public SystemAccessible FirstChild {
            get {
                try {
                    if (this.firstChild == null) {
                        this.firstChild = this.Navigate(AccessibleNavigation.FirstChild);                        
                    }
                    return (SystemAccessible)firstChild;
                } catch {
                    return null;
                }
            }
        }

        public SystemAccessible LastChild {
            get {
                try {
                    if (this.lastChild == null) {
                        this.lastChild = this.Navigate(AccessibleNavigation.LastChild);                        
                    }
                    return (SystemAccessible)lastChild;
                } catch {
                    return null;
                }
            }
        }

        public SystemAccessible NextSibling {
            get {
                try {
                    if (this.next == null) {
                        this.next = this.Navigate(AccessibleNavigation.Next);                        
                    }
                    return (SystemAccessible )next;
                } catch {
                    return null;
                }
            }
        }

        public SystemAccessible FindChild(string name) {
            return FindChild(name, new StringComparer());
        }

        public SystemAccessible[] GetChildren() {
            List<SystemAccessible> result = new List<SystemAccessible>();
            try {
                SystemAccessible child = this.FirstChild;
                if (child == null) {
                    for (int i = 1, n = this.GetChildCount(); i <= n; i++) {
                        child = (SystemAccessible)this.GetChild(i);
                        if (child != null) result.Add(child);
                    }
                } else {
                    SystemAccessible last = this.LastChild;
                    while (child != null) {
                        result.Add(child);
                        if (child == last)
                            break;
                        child = (SystemAccessible)child.Navigate(AccessibleNavigation.Next);
                    }
                }
            } catch {
            }
            return result.ToArray();
        }

        public SystemAccessible FindChild(string name, IComparer<string> ic) {
            foreach (SystemAccessible child in GetChildren()) {
                try {
                    string childName = child.Name;
#if DEBUG
                    Trace.WriteLine("Found child: " + childName);
#endif
                    if (ic.Compare(name, childName) == 0) {
                        return (SystemAccessible)child;
                    }
                } catch {
                }
            }
            return null;
        }

        public AccessibleObject FindChild(int i) {
            foreach (SystemAccessible child in GetChildren()) {
                i--;
                if (i == 0) {
                    return child;
                }
            }
            return null;
        }

        public AccessibleObject FindChildByRole(AccessibleRole role) {
            foreach (SystemAccessible child in GetChildren()) {
                if (child.Role == role) {
                    return child;
                }
            }
            return null;
        }

        public override AccessibleObject Parent {
            get {
                if (this.parent != null) return parent;
                return this.Navigate(AccessibleNavigation.Up);
            }
        }


        // Implementation --------------------------------------------------------
        SystemAccessible Wrap(object obj) {
            return Wrap(this, obj as IAccessible, CHILDID_SELF);
        }

        SystemAccessible Wrap(object obj, int id) {
            return Wrap(this, obj as IAccessible, id);
        }

        static SystemAccessible Wrap(AccessibleObject parent, IAccessible acc, int id) {
            if (acc == null) return null;
            if (s_Wrappers.ContainsKey(acc)) {
                return s_Wrappers[acc];
            }
            SystemAccessible wrapper = new SystemAccessible(parent, acc, id);
            s_Wrappers[acc] = wrapper;
            return wrapper;
        }

        const uint OBJID_CLIENT = 0xFFFFFFFC;
        static int CHILDID_SELF = 0; 

        public static SystemAccessible AccessibleObjectForWindow(IntPtr hwnd) {

            Guid guid = typeof(IAccessible).GUID;
            IntPtr ptr;
            int hr = AccessibleObjectFromWindow(hwnd, OBJID_CLIENT, ref guid, out ptr);
            if (hr == 0) {
                try {
                    object obj = Marshal.GetObjectForIUnknown(ptr);
                    return Wrap(null, obj as IAccessible, CHILDID_SELF);
                } finally {
                    Marshal.Release(ptr);
                }
            }
            throw new ApplicationException("AccessibleObjectFromWindow failed");
        }

        public static SystemAccessible AccessibleWindowObjectAt(Point pt) {
            POINT p = new POINT();
            p.x = pt.X;
            p.y = pt.Y;
            IntPtr hwnd = WindowFromPoint(p);
            if (hwnd != IntPtr.Zero) {
                return SystemAccessible.AccessibleObjectForWindow(hwnd);
            }
            return null;
        }

        public static SystemAccessible AccessibleObjectAt(Point pt) {
            POINT p = new POINT();
            p.x = pt.X;
            p.y = pt.Y;
            IAccessible acc;
            object id;
            int hr = AccessibleObjectFromPoint(p, out acc, out id);
            if (hr == 0) {
                return Wrap(null, acc, (int)id);
            }
            throw new ApplicationException("AccessibleObjectFromWindow failed");
        }

        [DllImport("OleAcc", CharSet = CharSet.Unicode)]
        static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint dwObjectID, ref Guid riid, out IntPtr pObject);

        [DllImport("OleAcc", CharSet = CharSet.Unicode)]
        static extern int WindowFromAccessibleObject(IAccessible pacc, out IntPtr hwnd);        

        [DllImport("User32")]
        static extern IntPtr WindowFromPoint(POINT pt);

        [DllImport("User32", CharSet = CharSet.Unicode)]
        static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

        [StructLayout(LayoutKind.Sequential)]
        struct POINT {
            public int x;
            public int y;
        };

        [DllImport("OleAcc", CharSet = CharSet.Unicode)]
        static extern int AccessibleObjectFromPoint(POINT ptScreen, out IAccessible acc, out object varChild);

    }

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int left;
        public int top;
        public int right;
        public int bottom;
    };
}
