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
        private tb.Entregador reg;

        public Form2()
        {
            InitializeComponent();
            entregadorDAO = new dao.EntregadorDAO();
            reg = entregadorDAO.GetUltimoEntregador();
            // cntrole1.Ultimo = true;
            Mostra(); 
        }

        private bool Mostra()
        {
            if (reg == null)
            {
                return false;
            } else
            {
                txtNome.Text = reg.Nome;
                txtTelefone.Text = reg.Telefone;
                return true;
            }
        }

        private void cntrole1_Load(object sender, EventArgs e)
        {

        }

        private void cntrole1_AcaoRealizada(object sender, AcaoEventArgs e)
        {
            switch (e.Acao)
            {
                case "ParaTras":
                    reg = entregadorDAO.ParaTraz();
                    if (!Mostra())
                    {
                        cntrole1.Ultimo = true;
                    }
                    break;
                case "ParaFrente":
                    reg = entregadorDAO.ParaFrente();
                    if (!Mostra())
                    {
                        cntrole1.Primeiro = true;
                    }
                    break;
                case "Editar":
                    //cntrole1.Editanto = true; ;
                    this.Text = "clicou";
                    break;

            }
        }
    }
}
