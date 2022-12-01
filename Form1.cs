using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpenModelFromServer
{
    public partial class Form1 : Form
    {
        public string FormResult;
        public Form1()
        {
            InitializeComponent();
            OKbutton.Click += OKbutton_Click;
            Cancelbutton.Click += Cancelbutton_Click;
        }

        private void OKbutton_Click(object sender, EventArgs e)
        {
            FormResult = textBox1.Text;
            this.Close();
        }

        private void Cancelbutton_Click(object sender, EventArgs e)
        {
            FormResult = null;
            this.Close();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
