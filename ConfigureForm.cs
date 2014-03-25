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
    public partial class ConfigureForm : Form
    {
        public ConfigureForm()
        {
            InitializeComponent();
        }

        protected virtual void OnApplyChanges(EventArgs e)
        {
            SettingsOwner parent = (SettingsOwner)this.Owner;

            int interval = Convert.ToInt32(timeIntervalInput.Text);
            int scaleFactor;

            int index = selectedTimeUnits.SelectedIndex;
            switch (index)
            {
                case 0:
                    scaleFactor = 1000;
                    break;
                case 1:
                    scaleFactor = 60 * 1000;
                    break;
                case 2:
                    scaleFactor = 60 * 60 * 1000;
                    break;
                case 3:
                    scaleFactor = 24 * 60 * 60 * 1000;
                    break;
                default:
                    scaleFactor = 1000;
                    break;
            }

            parent.Notify(interval * scaleFactor);
        }

        private void ConfigureForm_Load(object sender, EventArgs e)
        {
            SettingsOwner parent = (SettingsOwner)this.Owner;
            int interval = parent.interval;

            selectedTimeUnits.SelectedIndex = 0;
            timeIntervalInput.Text = Convert.ToString(interval/1000);

            applyButton.Enabled = false;
        }

        private void applyButton_Click(object sender, EventArgs e)
        {
            OnApplyChanges(e);
            applyButton.Enabled = false;
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            OnApplyChanges(e);
            this.DialogResult = DialogResult.OK;
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void aboutDesktopCuratorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (((SettingsOwner)this.Owner).parentContext.aboutForm == null || 
                    !((SettingsOwner)this.Owner).parentContext.aboutForm.Visible)
            {
                ((SettingsOwner)this.Owner).parentContext.aboutForm = new AboutForm();
                ((SettingsOwner)this.Owner).parentContext.aboutForm.ShowDialog(this);
            }
        }

        private void browseButton_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                SettingsOwner parent = (SettingsOwner)this.Owner;
                parent.Notify(folderBrowserDialog.SelectedPath);
            }

            applyButton.Enabled = true;
        }

        private void selectedTimeUnits_SelectedIndexChanged(object sender, EventArgs e)
        {
            applyButton.Enabled = true;
        }

        private void timeIntervalInput_TextChanged(object sender, EventArgs e)
        {
            applyButton.Enabled = true;
        }
    }
}
