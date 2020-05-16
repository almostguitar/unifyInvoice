using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UnifyInvoice
{
    public partial class frmParamUnifyInvoice : Form
    {
        public frmParamUnifyInvoice()
        {
            InitializeComponent();
        }

        public string FromCell { get; set; }
        public string ToCell { get; set; }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(this.textBox1.Text) && string.IsNullOrEmpty(this.textBox2.Text) && string.IsNullOrEmpty(txtColumna.Text))
            {
                this.DialogResult = DialogResult.OK;

                this.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}
