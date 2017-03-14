using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SpreadSheetConnector
{
    public partial class Form1 : Form
    {
        public List<UpdaterItem> items;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            items = UpdaterItem.DeserializeUpdaterItems();
            foreach (var item in items)
            {
                item.StartWatch();
            }
        }
    }
}
