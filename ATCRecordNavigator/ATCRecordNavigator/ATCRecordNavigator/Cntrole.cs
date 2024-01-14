using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ATCRecordNavigator
{
    public partial class Cntrole : UserControl
    {
        public event EventHandler<AcaoEventArgs> AcaoRealizada;
        private bool EmMudanca = false;

        private bool primeiro;
        public bool Primeiro
        {
            get { return primeiro; }
            set { primeiro = value; DecideBotoes(); }
        }

        private bool ultimo;
        public bool Ultimo
        {
            get { return ultimo; }
            set { ultimo = value; DecideBotoes(); }
        }

        private bool emEdicao;
        public bool EmEdicao
        {
            get { return emEdicao; }
            set { emEdicao = value; DecideBotoes(); }
        }

        private bool emAdicao;
        public bool EmAdicao
        {
            get { return emAdicao; }
            set { emAdicao = value; DecideBotoes(); }
        }

        public Cntrole()
        {
            InitializeComponent();
            this.Dock = DockStyle.Top;
        }

        private void DecideBotoes()
        {
            if (EmMudanca==false)
            {
                if (emAdicao)
                {
                    btnParaFrente.Enabled = false;
                    btnParaTras.Enabled = false;
                    btnApagar.Enabled = false;
                    btnPesquisa.Enabled = false;
                    btnAdicionar.Enabled = false;
                }
                else
                {
                    if (EmEdicao)
                    {
                        this.btnEditar.Visible = false;
                        this.btnApagar.Visible = false;
                        this.btnOk.Visible = true;
                        this.btnCancelar.Visible = true;
                    }
                    else
                    {
                        btnPesquisa.Enabled = true;
                        btnAdicionar.Enabled = true;
                        btnParaFrente.Enabled = !Primeiro;
                        btnParaTras.Enabled = !Ultimo;
                        btnEditar.Enabled = !EmEdicao && !EmAdicao;
                        btnApagar.Enabled = !EmEdicao && !EmAdicao;
                    }
                }                
            }
        }

        private void toolTip1_Popup(object sender, PopupEventArgs e)
        {

        }

        private void btnParaTras_Click_1(object sender, EventArgs e)
        {
            EmMudanca = true;
            primeiro = false;
            AcaoRealizada?.Invoke(this, new AcaoEventArgs("ParaTras"));
            EmMudanca = false;
            DecideBotoes();
        }

        private void btnParaFrente_Click(object sender, EventArgs e)
        {
            EmMudanca = true;
            ultimo = false;
            AcaoRealizada?.Invoke(this, new AcaoEventArgs("ParaFrente"));
            EmMudanca = false;
            DecideBotoes();
        }

        private void btnApagar_Click(object sender, EventArgs e)
        {
            AcaoRealizada?.Invoke(this, new AcaoEventArgs("Delete"));
        }

        private void btnEditar_Click(object sender, EventArgs e)
        {
            if (EmEdicao)
            {
                this.btnEditar.Visible = true;
                this.btnApagar.Visible = true;
                this.btnOk.Visible = false;
                this.btnCancelar.Visible = false;
                emEdicao = false;
            } else
            {
                this.btnEditar.Visible = false;
                this.btnApagar.Visible = false;
                this.btnOk.Visible = true;
                this.btnCancelar.Visible = true;
                emEdicao = true;
            }                    
            AcaoRealizada?.Invoke(this, new AcaoEventArgs("Editar"));
        }

        private void MostraEdicao()
        {
            this.btnEditar.Visible = true;
            this.btnApagar.Visible = true;
            this.btnOk.Visible = false;
            this.btnCancelar.Visible = false;
        }

        public void ResetarAparenciaEditar()
        {

        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            emEdicao = false;
            emAdicao = false;
            AcaoRealizada?.Invoke(this, new AcaoEventArgs("OK"));
            MostraEdicao();
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            emEdicao = false;
            emAdicao = false;
            AcaoRealizada?.Invoke(this, new AcaoEventArgs("CANC"));
            MostraEdicao();
        }
    }
}
