using BonifacioEntregas.tb;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace BonifacioEntregas
{
    public partial class fCadClientes : FormBase
    {

        List<CampoTagInfo> tagsDosCampos;

        public fCadClientes()
        {
            InitializeComponent();
            base.DAO = new dao.ClienteDAO();
            base.reg = (tb.IDataEntity)DAO.GetUltimo();
            base.Mostra();
            tagsDosCampos = base.LerTagsDosCamposDeTexto();
        }

        private void cntrole1_AcaoRealizada(object sender, AcaoEventArgs e)
        {
            switch (e.Acao)
            {
                case "Adicionar":
                    LimparCampos();
                    EmAdicao = true;
                    break;
                case "Delete":
                    reg = (tb.Cliente)DAO.Apagar(Direcao);
                    if (!Mostra())
                    {
                        if (Direcao == 1)
                        {
                            cntrole1.Ultimo = true;
                        }
                        else
                        {
                            cntrole1.Primeiro = true;
                        }
                    }
                    break;
                case "ParaTras":
                    Direcao = -1;
                    reg = (tb.Cliente)DAO.ParaTraz();
                    if (!Mostra())
                    {
                        cntrole1.Ultimo = true;
                    }
                    break;
                case "ParaFrente":
                    Direcao = 1; ;
                    reg = (tb.Cliente)DAO.ParaFrente();
                    if (!Mostra())
                    {
                        cntrole1.Primeiro = true;
                    }
                    break;
                case "Editar":
                    // this.Text = "clicou";
                    break;
                case "CANC":
                    reg = (tb.Cliente)DAO.GetEsse();
                    ResetarAparenciaControles();
                    Mostra();
                    break;
                case "OK":                    
                    Grava(tagsDosCampos);
                    break;
            }
        }

        private void fCadClientes_KeyUp(object sender, KeyEventArgs e)
        {
            base.cntrole1.EmEdicao = true;
        }

        private void cntrole1_Load(object sender, System.EventArgs e)
        {

        }


    }
}
