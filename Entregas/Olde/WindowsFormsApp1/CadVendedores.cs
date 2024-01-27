using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace BonifacioEntregas
{
    public partial class CadVendedores : BonifacioEntregas.FormBase
    {

        private tb.Vendedor clienteEspecifico;

        public CadVendedores()
        {
            InitializeComponent();
            base.DAO = new dao.VendedoresDAO();
            // /clienteEspecifico = DAO.GetUltimo() as tb.Vendedor;
            // base.reg = DAO.GetUltimo() as tb.Vendedor;
            base.reg = base.DAO.ge
                //(tb.Vendedor)DAO.GetUltimo();
            base.Mostra();
            base.LerTagsDosCamposDeTexto();
        }

        private void cntrole1_AcaoRealizada(object sender, AcaoEventArgs e)
        {
            base.cntrole1_AcaoRealizada(sender, e, clienteEspecifico);
        }

        private void CadVendedores_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                base.Cancela();
            }
            else
            {
                if (!base.Pesquisando)
                {
                    base.cntrole1.EmEdicao = true;
                }
            }
        }
    }
}
