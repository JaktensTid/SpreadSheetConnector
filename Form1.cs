using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading.Tasks;

namespace SpreadSheetConnector
{
    public partial class Form1 : Form
    {
        public List<UpdaterItem> items = UpdaterItem.DeserializeUpdaterItems();
        private GoogleConnector googleConnector;
        public Form1()
        {
            InitializeComponent();
            DataGridView.DataSource = items;
        }

        private void AddNewButton_Click(object sender, EventArgs e)
        {
            AddForm addForm = new AddForm(this);
            addForm.Show();
        }

        public void AppendNewUpdaterItem(UpdaterItem item)
        {
            items.Add(item);
            UpdaterItem.SerializeUpdaterItems(items);
            DataGridView.DataSource = items;
        }
        private void RemoveSelectedButton_Click(object sender, EventArgs e)
        {
            int selectedIndex = DataGridView.CurrentCell.RowIndex;
            items[selectedIndex].Dispose();
            items.RemoveAt(selectedIndex);
            DataGridView.DataSource = items;
        }

        private async void ConnectToGoogleButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.FilterIndex = 1;
            choofdlog.Multiselect = false;

            if (choofdlog.ShowDialog() == DialogResult.OK)
            {
                string sFileName = choofdlog.FileName;
                if(choofdlog.FileNames.Length > 0)
                {
                    googleConnector = new GoogleConnector();
                    await googleConnector.Connect(choofdlog.FileNames[0]);
                    if(googleConnector.Authorized)
                    {
                        ConnectionSuccessLabel.ForeColor = Color.Green;
                        ConnectionSuccessLabel.Text = "Connected";
                    }
                    else
                    {
                        MessageBox.Show("Something went wrong.");
                    }
                }           
            }
        }

    }
}
