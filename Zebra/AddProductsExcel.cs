using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Zebra
{
    public partial class AddProductsExcel : Form
    {
        public AddProductsExcel()
        {
            InitializeComponent();
        }

        private void AddProductsExcel_Load(object sender, EventArgs e)
        {

        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            Form1 form = new Form1();
            this.Hide();
            form.ShowDialog();
        }
    }
}
