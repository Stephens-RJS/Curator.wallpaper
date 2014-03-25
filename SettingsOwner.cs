using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Curator
{
    public partial class SettingsOwner : Form
    {
        private TrayIconApplicationContext _parentContext;
        private int _interval;
        private string _path;

        public SettingsOwner()
        {
            InitializeComponent();
        }

        public SettingsOwner(TrayIconApplicationContext parentContext)
        {
            this._parentContext = parentContext;
            this._interval = parentContext.FormTimer.Interval;
            InitializeComponent();
            this.Hide();
        }

        public void Notify(int interval)
        {
            this._interval = interval;
            parentContext.UpdateSettings();
        }

        public void Notify(string path)
        {
            this._path = path;
        }

        public string path { get { return _path; } }
        public int interval { get { return _interval; } }
        public TrayIconApplicationContext parentContext { get { return _parentContext; } }
    }
}
