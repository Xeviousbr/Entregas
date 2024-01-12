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
    public partial class Form2 : Form
    {
        private dao.EntregadorDAO entregadorDAO;

        public Form2()
        {
            InitializeComponent();
            entregadorDAO = new dao.EntregadorDAO();
            tb.Entregador reg = entregadorDAO.GetUltimoEntregador();
            Mostra(reg); 
        }

        private void Mostra(tb.Entregador reg)
        {
            txtNome.Text = reg.Nome;
            txtTelefone.Text = reg.Telefone;
        }
    }
}
