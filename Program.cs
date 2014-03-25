using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.Win32;

using System.Diagnostics;

namespace Curator
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new TrayIconApplicationContext());
        }
    }

    public class TrayIconApplicationContext : ApplicationContext
    {
        private readonly NotifyIcon _notifyIcon;
        private readonly ContextMenuStrip _contextMenu;
        private readonly System.Windows.Forms.Timer _timer;

        private SettingsOwner _settingsOwner;
        private WallpaperChanger _wallpaperChanger;
        private ConfigureForm _configureInstance = null;
        private AboutForm _aboutInstance = null;

        public TrayIconApplicationContext()
        {
            _contextMenu = new ContextMenuStrip();
            this.ContextMenu.Items.Add("&Configure", null, this.ConfigureContextMenuClickHandler).Font = new Font(this.ContextMenu.Font, FontStyle.Bold);
            this.ContextMenu.Items.Add("&Next wallpaper", null, this.NextContextMenuClickHandler);
            this.ContextMenu.Items.Add("-");
            this.ContextMenu.Items.Add("&About", null, this.AboutContextMenuClickHandler);
            this.ContextMenu.Items.Add("-");
            this.ContextMenu.Items.Add("E&xit", null, this.ExitContextMenuClickHandler);

            Application.ApplicationExit += this.ApplicationExitHandler;

            _notifyIcon = new NotifyIcon
            {
                Text = "Desktop Curator",

                // Needs to be fixed. Add .ico's as resources?
                Icon = new Icon(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal)
                       + @"\TrayIcon.ico"),

                ContextMenuStrip = _contextMenu,
                Visible = true
            };

            _timer = new System.Windows.Forms.Timer
            {
                Enabled = true,
                Interval = 30000
            };
            _timer.Tick += new EventHandler(timer_Tick);

            _settingsOwner = new SettingsOwner(this);
            _wallpaperChanger = new WallpaperChanger();

            this.TrayIcon.MouseDoubleClick += this.TrayIconDoubleClickHandler;
            this.TrayIcon.MouseClick += this.TrayIconClickHandler;

            this.ShowConfiguration();
            _timer.Start();
        }

        public ContextMenuStrip ContextMenu
        {
            get { return _contextMenu; }
        }

        public NotifyIcon TrayIcon
        {
            get { return _notifyIcon; }
        }

        public System.Windows.Forms.Timer FormTimer
        {
            get { return _timer;  }
        }

        public AboutForm aboutForm
        {
            get { return _aboutInstance; }
            set { _aboutInstance = value; }
        }

        protected virtual void timer_Tick(object sender, EventArgs e)
        {
            _wallpaperChanger.ChangeWallpaper();
        }

        protected virtual void OnTrayIconClick(MouseEventArgs e)
        { }

        protected virtual void OnTrayIconDoubleClick(MouseEventArgs e)
        {
            this.ShowConfiguration();
        }

        protected virtual void OnApplicationExit(EventArgs e)
        {
            if (_contextMenu != null)
                _contextMenu.Dispose();

            if (_timer != null)
                _timer.Dispose();

            if (_notifyIcon != null)
            {
                _notifyIcon.Visible = false;
                _notifyIcon.Dispose();
            }
        }

        private void TrayIconClickHandler(object sender, MouseEventArgs e)
        {
            this.OnTrayIconClick(e);
        }

        private void TrayIconDoubleClickHandler(object sender, MouseEventArgs e)
        {
            this.OnTrayIconDoubleClick(e);
        }

        private void NextContextMenuClickHandler(object sender, EventArgs eventArgs)
        {
            this._timer.Stop();
            _wallpaperChanger.ChangeWallpaper();
            this._timer.Start();
        }

        private void AboutContextMenuClickHandler(object sender, EventArgs eventArgs)
        {
            if (this._aboutInstance == null || !this._aboutInstance.Visible)
            {
                _aboutInstance = new AboutForm();
                _aboutInstance.ShowDialog(this._configureInstance);

                this._configureInstance.TopMost = true;
                this._configureInstance.TopMost = false;

                foreach (Form form in this._configureInstance.OwnedForms)
                {
                    form.TopMost = true;
                    form.TopMost = false;
                }
            }
            else
            {
                Curator.WallpaperChanger.FlashWindow.Flash(aboutForm, 4);
            }

            this._aboutInstance.TopMost = true;
            this._aboutInstance.TopMost = false;            
        }
        
        private void ConfigureContextMenuClickHandler(object sender, EventArgs eventArgs)
        {
            if (this._aboutInstance == null || !this._aboutInstance.Visible)
            {
                this.ShowConfiguration();
            }            
            else
            {
                Curator.WallpaperChanger.FlashWindow.Flash(aboutForm, 4);
                this._aboutInstance.TopMost = true;
                this._aboutInstance.TopMost = false;
            }
        }

        private void ExitContextMenuClickHandler(object sender, EventArgs eventArgs)
        {
            Application.Exit();
        }

        private void ApplicationExitHandler(object sender, EventArgs e)
        {
            this.OnApplicationExit(e);
        }

        private void ShowConfiguration()
        {
            if (this._configureInstance == null || !this._configureInstance.Visible)
            {
                this._configureInstance = new ConfigureForm();
                this._configureInstance.ShowDialog(_settingsOwner);
            }
            else
            {
                Curator.WallpaperChanger.FlashWindow.Flash(aboutForm, 4);
                this._configureInstance.TopMost = true;
                this._configureInstance.TopMost = false;
            }
                    
            if (this._configureInstance.OwnedForms.Length > 0)
            {
                foreach (Form form in this._configureInstance.OwnedForms)
                {
                    form.TopMost = true;
                    form.TopMost = false;
                }
            }
        }

        public void UpdateSettings()
        {
            this._timer.Interval = this._settingsOwner.interval;
            this._wallpaperChanger.path = this._settingsOwner.path;
        }
    }

    class WallpaperChanger
    {
        private string _path;
        static Random randGen; 

        public WallpaperChanger()
        {
            _path = null;
            randGen = new Random();
        }

        public string path { get { return _path; } set { _path = value; } }

        public void ChangeWallpaper()
        {
            if (path != null)
            {
                List<string> images = new List<string>();
                var filters = new String[] { "jpg", "jpeg", "png", "gif", "tiff", "bmp" };
                foreach (var filter in filters)
                {
                    images.AddRange(System.IO.Directory.GetFiles(path, String.Format("*.{0}", filter), System.IO.SearchOption.AllDirectories));
                }

                string fileName = images[randGen.Next(0, images.Count)];

                // Temporary resizing method... essentially just preserving quality with Fit

                float width = Screen.PrimaryScreen.Bounds.Width;
                float height = Screen.PrimaryScreen.Bounds.Height;
                var brush = new SolidBrush(Color.Black);

                Bitmap image = new Bitmap(fileName);
                Bitmap bmp = new Bitmap((int)width, (int)height);
                float scale = Math.Min(width / image.Width, height / image.Height);
                var graph = Graphics.FromImage(bmp);

                graph.InterpolationMode = InterpolationMode.High;
                graph.CompositingQuality = CompositingQuality.HighQuality;
                graph.SmoothingMode = SmoothingMode.AntiAlias;

                var scaleWidth = (int)(image.Width * scale);
                var scaleHeight = (int)(image.Height * scale);

                graph.FillRectangle(brush, new RectangleF(0, 0, width, height));
                graph.DrawImage(image, new Rectangle(((int)width - scaleWidth) / 2, ((int)height - scaleHeight) / 2, scaleWidth, scaleHeight));
                graph.Dispose();

                if (!Directory.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Curator\temp")))
                {
                    Directory.CreateDirectory(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Curator\temp"));
                }

                try
                {
                    bmp.Save(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Curator\temp\temp.bmp"), System.Drawing.Imaging.ImageFormat.Bmp);
                }
                catch (Exception e)
                {
                    Debug.WriteLine("DEBUG::ChangeWallpaper()::Error attempting to save temporary image::" + e.Message);
                }

                // Clean up clean up everybody everywhere
                image.Dispose();
                bmp.Dispose();

                IntPtr result = IntPtr.Zero;
                WinAPI.SendMessageTimeout(WinAPI.FindWindow("Progman", null), 0x52c, IntPtr.Zero, IntPtr.Zero, 0, 500, out result);

                ThreadStart threadStarter = () =>
                {
                    WinAPI.IActiveDesktop _activeDesktop = WinAPI.ActiveDesktopWrapper.GetActiveDesktop();
                    _activeDesktop.SetWallpaper(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Curator\temp\temp.bmp"), 0);
                    _activeDesktop.ApplyChanges(WinAPI.AD_Apply.ALL | WinAPI.AD_Apply.FORCE);

                    Marshal.ReleaseComObject(_activeDesktop);
                };
                Thread thread = new Thread(threadStarter);
                thread.SetApartmentState(ApartmentState.STA); //Set the thread to STA (REQUIRED!!!!)
                thread.Start();
                thread.Join(2000);
            }
        }

        public static class FlashWindow
        {
            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            private static extern bool FlashWindowEx(ref FLASHWINFO pwfi);

            [StructLayout(LayoutKind.Sequential)]
            private struct FLASHWINFO
            {
                /// <summary>
                /// The size of the structure in bytes.
                /// </summary>
                public uint cbSize;
                /// <summary>
                /// A Handle to the Window to be Flashed. The window can be either opened or minimized.
                /// </summary>
                public IntPtr hwnd;
                /// <summary>
                /// The Flash Status.
                /// </summary>
                public uint dwFlags;
                /// <summary>
                /// The number of times to Flash the window.
                /// </summary>
                public uint uCount;
                /// <summary>
                /// The rate at which the Window is to be flashed, in milliseconds. If Zero, the function uses the default cursor blink rate.
                /// </summary>
                public uint dwTimeout;
            }

            /// <summary>
            /// Stop flashing. The system restores the window to its original stae.
            /// </summary>
            public const uint FLASHW_STOP = 0;

            /// <summary>
            /// Flash the window caption.
            /// </summary>
            public const uint FLASHW_CAPTION = 1;

            /// <summary>
            /// Flash the taskbar button.
            /// </summary>
            public const uint FLASHW_TRAY = 2;

            /// <summary>
            /// Flash both the window caption and taskbar button.
            /// This is equivalent to setting the FLASHW_CAPTION | FLASHW_TRAY flags.
            /// </summary>
            public const uint FLASHW_ALL = 3;

            /// <summary>
            /// Flash continuously, until the FLASHW_STOP flag is set.
            /// </summary>
            public const uint FLASHW_TIMER = 4;

            /// <summary>
            /// Flash continuously until the window comes to the foreground.
            /// </summary>
            public const uint FLASHW_TIMERNOFG = 12;


            /// <summary>
            /// Flash the spacified Window (Form) until it recieves focus.
            /// </summary>
            /// <param name="form">The Form (Window) to Flash.</param>
            /// <returns></returns>
            public static bool Flash(System.Windows.Forms.Form form)
            {
                // Make sure we're running under Windows 2000 or later
                if (Win2000OrLater)
                {
                    FLASHWINFO fi = Create_FLASHWINFO(form.Handle, FLASHW_ALL | FLASHW_TIMERNOFG, uint.MaxValue, 0);
                    return FlashWindowEx(ref fi);
                }
                return false;
            }

            private static FLASHWINFO Create_FLASHWINFO(IntPtr handle, uint flags, uint count, uint timeout)
            {
                FLASHWINFO fi = new FLASHWINFO();
                fi.cbSize = Convert.ToUInt32(Marshal.SizeOf(fi));
                fi.hwnd = handle;
                fi.dwFlags = flags;
                fi.uCount = count;
                fi.dwTimeout = timeout;
                return fi;
            }

            /// <summary>
            /// Flash the specified Window (form) for the specified number of times
            /// </summary>
            /// <param name="form">The Form (Window) to Flash.</param>
            /// <param name="count">The number of times to Flash.</param>
            /// <returns></returns>
            public static bool Flash(System.Windows.Forms.Form form, uint count)
            {
                if (Win2000OrLater)
                {
                    FLASHWINFO fi = Create_FLASHWINFO(form.Handle, FLASHW_ALL, count, 0);
                    return FlashWindowEx(ref fi);
                }
                return false;
            }

            /// <summary>
            /// Start Flashing the specified Window (form)
            /// </summary>
            /// <param name="form">The Form (Window) to Flash.</param>
            /// <returns></returns>
            public static bool Start(System.Windows.Forms.Form form)
            {
                if (Win2000OrLater)
                {
                    FLASHWINFO fi = Create_FLASHWINFO(form.Handle, FLASHW_ALL, uint.MaxValue, 0);
                    return FlashWindowEx(ref fi);
                }
                return false;
            }

            /// <summary>
            /// Stop Flashing the specified Window (form)
            /// </summary>
            /// <param name="form"></param>
            /// <returns></returns>
            public static bool Stop(System.Windows.Forms.Form form)
            {
                if (Win2000OrLater)
                {
                    FLASHWINFO fi = Create_FLASHWINFO(form.Handle, FLASHW_STOP, uint.MaxValue, 0);
                    return FlashWindowEx(ref fi);
                }
                return false;
            }

            /// <summary>
            /// A boolean value indicating whether the application is running on Windows 2000 or later.
            /// </summary>
            private static bool Win2000OrLater
            {
                get { return System.Environment.OSVersion.Version.Major >= 5; }
            }
        }       

        public static class WinAPI
        {
            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
            public const int SPI_SETDESKWALLPAPER = 20;
            public const int SPIF_SENDCHANGE = 0x2;

            [DllImport("user32.dll", SetLastError = true)]
            public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            public static extern int SendMessageTimeout(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam, uint fuFlags, uint uTimeout, out IntPtr result);

            [Flags]
            public enum AD_Apply : int
            {
                SAVE = 0x00000001,
                HTMLGEN = 0x00000002,
                REFRESH = 0x00000004,
                ALL = SAVE | HTMLGEN | REFRESH,
                FORCE = 0x00000008,
                BUFFERED_REFRESH = 0x00000010,
                DYNAMICREFRESH = 0x00000020
            }

            [StructLayout(LayoutKind.Sequential)]
            public struct WALLPAPEROPT
            {
                public static readonly int SizeOf = Marshal.SizeOf(typeof(WALLPAPEROPT));
                public WallPaperStyle dwStyle;
            }

            public enum WallPaperStyle : int
            {
                WPSTYLE_CENTER = 0,
                WPSTYLE_TILE = 1,
                WPSTYLE_STRETCH = 2,
                /// <summary>Introduced in Windows 7</summary>
                WPSTYLE_KEEPASPECT = 3,
                /// <summary>Introduced in Windows 7</summary>
                WPSTYLE_CROPTOFIT = 4,
                /// <summary>Introduced in Windows 8</summary>
                WPSTYLE_SPAN = 5,
                WPSTYLE_MAX = 5
            }

            [StructLayout(LayoutKind.Sequential)]
            public struct COMPONENTSOPT
            {
                public static readonly int SizeOf = Marshal.SizeOf(typeof(COMPONENTSOPT));
                public int dwSize;
                [MarshalAs(UnmanagedType.Bool)]
                public bool fEnableComponents;
                [MarshalAs(UnmanagedType.Bool)]
                public bool fActiveDesktop;
            }

            [Flags]
            public enum CompItemState : int
            {
                NORMAL = 0x00000001,
                FULLSCREEN = 00000002,
                SPLIT = 0x00000004,
                VALIDSIZESTATEBITS = NORMAL | SPLIT | FULLSCREEN,
                VALIDSTATEBITS = NORMAL | SPLIT | FULLSCREEN | unchecked((int)0x80000000) | 0x40000000
            }

            [StructLayout(LayoutKind.Sequential)]
            public struct COMPSTATEINFO
            {
                public static readonly int SizeOf = Marshal.SizeOf(typeof(COMPSTATEINFO));
                public int dwSize;
                public int iLeft;
                public int iTop;
                public int dwWidth;
                public int dwHeight;
                public CompItemState dwItemState;
            }

            [StructLayout(LayoutKind.Sequential)]
            public struct COMPPOS
            {
                public const int COMPONENT_TOP = 0x3FFFFFFF;
                public const int COMPONENT_DEFAULT_LEFT = 0xFFFF;
                public const int COMPONENT_DEFAULT_TOP = 0xFFFF;
                public static readonly int SizeOf = Marshal.SizeOf(typeof(COMPPOS));

                public int dwSize;
                public int iLeft;
                public int iTop;
                public int dwWidth;
                public int dwHeight;
                public int izIndex;
                [MarshalAs(UnmanagedType.Bool)]
                public bool fCanResize;
                [MarshalAs(UnmanagedType.Bool)]
                public bool fCanResizeX;
                [MarshalAs(UnmanagedType.Bool)]
                public bool fCanResizeY;
                public int iPreferredLeftPercent;
                public int iPreferredTopPercent;
            }

            public enum CompType : int
            {
                HTMLDOC = 0,
                PICTURE = 1,
                WEBSITE = 2,
                CONTROL = 3,
                CFHTML = 4
            }

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode, Pack = 8)]
            public struct COMPONENT
            {
                private const int INTERNET_MAX_URL_LENGTH = 2084; // = INTERNET_MAX_SCHEME_LENGTH (32) + "://\0".Length +   INTERNET_MAX_PATH_LENGTH (2048)
                public static readonly int SizeOf = Marshal.SizeOf(typeof(COMPONENT));

                public int dwSize;
                public int dwID;
                public CompType iComponentType;
                [MarshalAs(UnmanagedType.Bool)]
                public bool fChecked;
                [MarshalAs(UnmanagedType.Bool)]
                public bool fDirty;
                [MarshalAs(UnmanagedType.Bool)]
                public bool fNoScroll;
                public COMPPOS cpPos;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
                public string wszFriendlyName;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = INTERNET_MAX_URL_LENGTH)]
                public string wszSource;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = INTERNET_MAX_URL_LENGTH)]
                public string wszSubscribedURL;

                public int dwCurItemState;
                public COMPSTATEINFO csiOriginal;
                public COMPSTATEINFO csiRestored;
            }

            public enum DtiAddUI : int
            {
                DEFAULT = 0x00000000,
                DISPSUBWIZARD = 0x00000001,
                POSITIONITEM = 0x00000002,
            }

            [Flags]
            public enum ComponentModify : int
            {
                TYPE = 0x00000001,
                CHECKED = 0x00000002,
                DIRTY = 0x00000004,
                NOSCROLL = 0x00000008,
                POS_LEFT = 0x00000010,
                POS_TOP = 0x00000020,
                SIZE_WIDTH = 0x00000040,
                SIZE_HEIGHT = 0x00000080,
                POS_ZINDEX = 0x00000100,
                SOURCE = 0x00000200,
                FRIENDLYNAME = 0x00000400,
                SUBSCRIBEDURL = 0x00000800,
                ORIGINAL_CSI = 0x00001000,
                RESTORED_CSI = 0x00002000,
                CURITEMSTATE = 0x00004000,
                ALL = TYPE | CHECKED | DIRTY | NOSCROLL | POS_LEFT | SIZE_WIDTH |
                    SIZE_HEIGHT | POS_ZINDEX | SOURCE |
                    FRIENDLYNAME | POS_TOP | SUBSCRIBEDURL | ORIGINAL_CSI |
                    RESTORED_CSI | CURITEMSTATE
            }

            [Flags]
            public enum AddURL : int
            {
                SILENT = 0x0001
            }

            [ComImport]
            [Guid("F490EB00-1240-11D1-9888-006097DEACF9")]
            [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            public interface IActiveDesktop
            {
                [PreserveSig]
                int ApplyChanges(AD_Apply dwFlags);
                [PreserveSig]
                int GetWallpaper([MarshalAs(UnmanagedType.LPWStr)]  System.Text.StringBuilder pwszWallpaper,
                          int cchWallpaper,
                          int dwReserved);
                [PreserveSig]
                int SetWallpaper([MarshalAs(UnmanagedType.LPWStr)] string pwszWallpaper, int dwReserved);
                [PreserveSig]
                int GetWallpaperOptions(ref WALLPAPEROPT pwpo, int dwReserved);
                [PreserveSig]
                int SetWallpaperOptions(ref WALLPAPEROPT pwpo, int dwReserved);
                [PreserveSig]
                int GetPattern([MarshalAs(UnmanagedType.LPWStr)] System.Text.StringBuilder pwszPattern, int cchPattern, int dwReserved);
                [PreserveSig]
                int SetPattern([MarshalAs(UnmanagedType.LPWStr)] string pwszPattern, int dwReserved);
                [PreserveSig]
                int GetDesktopItemOptions(ref COMPONENTSOPT pco, int dwReserved);
                [PreserveSig]
                int SetDesktopItemOptions(ref COMPONENTSOPT pco, int dwReserved);
                [PreserveSig]
                int AddDesktopItem(ref COMPONENT pcomp, int dwReserved);
                [PreserveSig]
                int AddDesktopItemWithUI(IntPtr hwnd, ref COMPONENT pcomp, DtiAddUI dwFlags);
                [PreserveSig]
                int ModifyDesktopItem(ref COMPONENT pcomp, ComponentModify dwFlags);
                [PreserveSig]
                int RemoveDesktopItem(ref COMPONENT pcomp, int dwReserved);
                [PreserveSig]
                int GetDesktopItemCount(out int lpiCount, int dwReserved);
                [PreserveSig]
                int GetDesktopItem(int nComponent, ref COMPONENT pcomp, int dwReserved);
                [PreserveSig]
                int GetDesktopItemByID(IntPtr dwID, ref COMPONENT pcomp, int dwReserved);
                [PreserveSig]
                int GenerateDesktopItemHtml([MarshalAs(UnmanagedType.LPWStr)] string pwszFileName, ref COMPONENT pcomp, int dwReserved);
                [PreserveSig]
                int AddUrl(IntPtr hwnd, [MarshalAs(UnmanagedType.LPWStr)] string pszSource, ref COMPONENT pcomp, AddURL dwFlags);
                [PreserveSig]
                int GetDesktopItemBySource([MarshalAs(UnmanagedType.LPWStr)] string pwszSource, ref COMPONENT pcomp, int dwReserved);
            }

            public class ActiveDesktopWrapper
            {
                static readonly Guid CLSID_ActiveDesktop = new Guid("{75048700-EF1F-11D0-9888-006097DEACF9}");

                public static IActiveDesktop GetActiveDesktop()
                {
                    Type typeActiveDesktop = Type.GetTypeFromCLSID(CLSID_ActiveDesktop);
                    return (IActiveDesktop)Activator.CreateInstance(typeActiveDesktop);
                }
            }                        
        } 
    }
}
