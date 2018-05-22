using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using testeExcel;

namespace testeCampos
{
    public partial class Form2 : Form
    {
        private Form1 formPrincipal;

        public Form2(Form1 formPrincipal) : this()
        {
            this.formPrincipal = formPrincipal;
        }

        
        public Form2()
        {
            InitializeComponent();
 
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            formPrincipal.checado = true;
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            formPrincipal.checado = false;
        }
    }
}
