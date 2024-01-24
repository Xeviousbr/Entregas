using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BonifacioEntregas
{
    public partial class oprConfig : Form
    {
        public oprConfig()
        {
            InitializeComponent();
            textBox1.Text = gen.CaminhoBase;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = (openFileDialog1.FileName);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            INI MeuIni = new INI();
            gen.CaminhoBase = textBox1.Text;
            MeuIni.WriteString("Config", "Base", gen.CaminhoBase);
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
