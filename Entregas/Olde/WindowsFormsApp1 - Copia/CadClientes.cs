using BonifacioEntregas.tb;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using SourceGrid;
using System.Data;

namespace BonifacioEntregas
{
    public partial class fCadClientes : FormBase
    {

        private tb.Cliente clienteEspecifico;
        private bool txtIdDentro = false;

        public fCadClientes()
        {
            InitializeComponent();
            base.DAO = new dao.ClienteDAO();
            clienteEspecifico = DAO.GetUltimo() as tb.Cliente;
            base.reg = DAO.GetUltimo() as tb.Cliente;
            base.Mostra();
            base.LerTagsDosCamposDeTexto();
            //base.NrLinhas(100);
        }

        private void cntrole1_AcaoRealizada(object sender, AcaoEventArgs e)
        {
            base.cntrole1_AcaoRealizada(sender, e, clienteEspecifico);
        }

        private void fCadClientes_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                base.Cancela();
            }
            else
            {
                if (!base.Pesquisando)
                {
                    if (!this.txtIdDentro)
                    {
                        base.cntrole1.EmEdicao = true;
                    }                    
                }                
            }
        }

        private void cntrole1_Load(object sender, System.EventArgs e)
        {

        }

        private void txtTelefone_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && !char.Equals(e.KeyChar, '-') && !char.Equals(e.KeyChar, '/') && !char.Equals(e.KeyChar, '(') && !char.Equals(e.KeyChar, ')'))
            {
                e.Handled = true;
            }
        }

        private void fCadClientes_Activated(object sender, EventArgs e)
        {
        }

        private void txtId_Enter(object sender, EventArgs e)
        {
            this.txtIdDentro = true;
        }

        private void txtId_Leave(object sender, EventArgs e)
        {
            this.txtIdDentro = false;
        }
    }
}
