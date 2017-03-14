using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SpreadSheetConnector
{
    public partial class AddForm : Form
    {
        private Form caller;
        private string name;
        private string folder;
        private string googleUrl;
        private UpdaterAction action;
        private Tuple<int, int> removableRows;

        public AddForm(Form caller)
        {
            InitializeComponent();
            this.caller = caller;
        }

        private void AddButton_Click(object sender, EventArgs e)
        {
            Action<string> showMessage = (message) => { MessageBox.Show(message); };
            name = NameTextBox.Text;
            googleUrl = GoogleSheetPathTextBox.Text;
            action = ActionComboBox.SelectedText == "Overwrite" ? UpdaterAction.Overwrite : UpdaterAction.Append;
            int fromRange;
            int toRange;
            bool t1 = int.TryParse(FromRangeTextBox.Text, out fromRange);
            bool t2 = int.TryParse(ToRangeTextBox.Text, out toRange);
            if (name == "" || name == "Name") { showMessage("Enter name"); return; }
            if (googleUrl == "" || googleUrl == "Google url") { showMessage("Enter google url"); return; }
            if (folder == "") { showMessage("Pick folder"); return; }
            if (action == default(UpdaterAction)) { showMessage("Select action"); return; }
            if (!t1 | !t2) removableRows = null; else removableRows = new Tuple<int, int>(fromRange, toRange);
            (caller as Form1).AppendNewUpdaterItem(new UpdaterItem(name, folder, googleUrl, action, removableRows));
            Close();
        }

        private void PickFolderButton_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                folder = folderBrowserDialog.SelectedPath;
            }
        }
    }
}
