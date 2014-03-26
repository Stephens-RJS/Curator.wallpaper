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
            get { return _timer; }
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
                WinAPI.FlashWindow.Flash(aboutForm, 4);
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
                WinAPI.FlashWindow.Flash(aboutForm, 4);
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
                WinAPI.FlashWindow.Flash(aboutForm, 4);
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
            //this._timer.Stop();
            this._timer.Interval = this._settingsOwner.interval;
            this._wallpaperChanger.path = this._settingsOwner.path;
            //this._timer.Start();
        }
    }
}
