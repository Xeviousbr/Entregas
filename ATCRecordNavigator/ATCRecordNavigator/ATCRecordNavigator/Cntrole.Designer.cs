
using System;
using System.Drawing;
using System.Windows.Forms;

namespace ATCRecordNavigator
{
    partial class Cntrole
    {
        /// <summary> 
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Limpar os recursos que estão sendo usados.
        /// </summary>
        /// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código gerado pelo Designer de Componentes

        /// <summary> 
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Cntrole));
            this.txtPesquisar = new System.Windows.Forms.ToolStripTextBox();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.btnPesquisa = new System.Windows.Forms.ToolStripButton();
            this.btnAdicionar = new System.Windows.Forms.ToolStripButton();
            this.btnOk = new System.Windows.Forms.ToolStripButton();
            this.btnCancelar = new System.Windows.Forms.ToolStripButton();
            this.btnEditar = new System.Windows.Forms.ToolStripButton();
            this.btnApagar = new System.Windows.Forms.ToolStripButton();
            this.btnParaTras = new System.Windows.Forms.ToolStripButton();
            this.btnParaFrente = new System.Windows.Forms.ToolStripButton();
            this.pesquisaTimer = new System.Windows.Forms.Timer(this.components);
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtPesquisar
            // 
            this.txtPesquisar.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtPesquisar.Name = "txtPesquisar";
            this.txtPesquisar.Size = new System.Drawing.Size(100, 23);
            this.txtPesquisar.Visible = false;
            this.txtPesquisar.KeyUp += new System.Windows.Forms.KeyEventHandler(this.txtPesquisar_KeyUp);
            // 
            // toolStrip1
            // 
            this.toolStrip1.ImageScalingSize = new System.Drawing.Size(48, 48);
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnPesquisa,
            this.btnAdicionar,
            this.btnOk,
            this.btnCancelar,
            this.btnEditar,
            this.btnApagar,
            this.btnParaTras,
            this.btnParaFrente,
            this.txtPesquisar});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(516, 55);
            this.toolStrip1.TabIndex = 0;
            // 
            // btnPesquisa
            // 
            this.btnPesquisa.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnPesquisa.Image = ((System.Drawing.Image)(resources.GetObject("btnPesquisa.Image")));
            this.btnPesquisa.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnPesquisa.Name = "btnPesquisa";
            this.btnPesquisa.Size = new System.Drawing.Size(52, 52);
            this.btnPesquisa.Text = "Pesquisa";
            this.btnPesquisa.ToolTipText = "Pesquisa";
            this.btnPesquisa.Click += new System.EventHandler(this.btnPesquisa_Click);
            // 
            // btnAdicionar
            // 
            this.btnAdicionar.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.btnAdicionar.Image = ((System.Drawing.Image)(resources.GetObject("btnAdicionar.Image")));
            this.btnAdicionar.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnAdicionar.Name = "btnAdicionar";
            this.btnAdicionar.Size = new System.Drawing.Size(52, 52);
            this.btnAdicionar.Text = "Adicionar";
            this.btnAdicionar.Click += new System.EventHandler(this.btnAdicionar_Click);
            // 
            // btnOk
            // 
            this.btnOk.Image = ((System.Drawing.Image)(resources.GetObject("btnOk.Image")));
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(52, 52);
            this.btnOk.Visible = false;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // btnCancelar
            // 
            this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
            this.btnCancelar.Name = "btnCancelar";
            this.btnCancelar.Size = new System.Drawing.Size(52, 52);
            this.btnCancelar.Visible = false;
            this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
            // 
            // btnEditar
            // 
            this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
            this.btnEditar.Name = "btnEditar";
            this.btnEditar.Size = new System.Drawing.Size(52, 52);
            this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
            // 
            // btnApagar
            // 
            this.btnApagar.Image = ((System.Drawing.Image)(resources.GetObject("btnApagar.Image")));
            this.btnApagar.Name = "btnApagar";
            this.btnApagar.Size = new System.Drawing.Size(52, 52);
            this.btnApagar.Click += new System.EventHandler(this.btnApagar_Click);
            // 
            // btnParaTras
            // 
            this.btnParaTras.Image = ((System.Drawing.Image)(resources.GetObject("btnParaTras.Image")));
            this.btnParaTras.Name = "btnParaTras";
            this.btnParaTras.Size = new System.Drawing.Size(52, 52);
            this.btnParaTras.Click += new System.EventHandler(this.btnParaTras_Click_1);
            // 
            // btnParaFrente
            // 
            this.btnParaFrente.Image = ((System.Drawing.Image)(resources.GetObject("btnParaFrente.Image")));
            this.btnParaFrente.Name = "btnParaFrente";
            this.btnParaFrente.Size = new System.Drawing.Size(52, 52);
            this.btnParaFrente.Click += new System.EventHandler(this.btnParaFrente_Click);
            // 
            // pesquisaTimer
            // 
            this.pesquisaTimer.Interval = 500;
            this.pesquisaTimer.Tick += new System.EventHandler(this.pesquisaTimer_Tick);
            // 
            // Cntrole
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.toolStrip1);
            this.Name = "Cntrole";
            this.Size = new System.Drawing.Size(516, 54);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void btnAdicionar_Click(object sender, EventArgs e)
        {
            emAdicao = true;
            AcaoRealizada?.Invoke(this, new AcaoEventArgs("Adicionar"));
            MostraEmEstadodeEdicao();
        }

        //private void btnPesquisarAdicional_Click(object sender, EventArgs e)
        //{
        //    this.btnPesquisarAdicional.Visible = false;
        //    this.txtPesquisar.Visible = false;
        //    this.btnParaFrente.Visible = true;
        //    this.btnParaTras.Visible = true;
        //    this.btnEditar.Visible = true;
        //    this.btnApagar.Visible = true;
        //    AcaoRealizada?.Invoke(this, new AcaoEventArgs("PesqAcionar"));
        //}

        private void btnPesquisa_Click(object sender, EventArgs e)
        {
            bool isSearchVisible = this.txtPesquisar.Visible;
            this.btnParaFrente.Visible = isSearchVisible;
            this.btnParaTras.Visible = isSearchVisible;
            this.btnEditar.Visible = isSearchVisible;
            this.btnApagar.Visible = isSearchVisible;
            this.btnAdicionar.Visible = isSearchVisible;
            this.txtPesquisar.Visible = !isSearchVisible;
            if (this.txtPesquisar.Visible)
            {
                this.txtPesquisar.AutoSize = false;
                int tamanho = this.toolStrip1.Width - 70;
                this.txtPesquisar.Size = new Size(tamanho, this.txtPesquisar.Height);
                this.txtPesquisar.Font = new Font(this.txtPesquisar.Font.FontFamily, 12);
                this.txtPesquisar.Focus();
                AcaoRealizada?.Invoke(this, new AcaoEventArgs("PesqON"));
            }
            else
            {
                this.txtPesquisar.AutoSize = true;
                this.txtPesquisar.Font = new Font(this.txtPesquisar.Font.FontFamily, 8.25F);
                this.btnAdicionar.Visible = true;
            }
        }

        #endregion

        private System.Windows.Forms.ToolTip toolTip1;
        private ToolStrip toolStrip1;
        private ToolStripButton btnPesquisa;        
        private ToolStripButton btnParaTras;
        private ToolStripButton btnParaFrente;
        private ToolStripButton btnEditar;
        private ToolStripButton btnApagar;
        private ToolStripTextBox txtPesquisar;
        private System.Windows.Forms.ToolStripButton btnAdicionar;
        private System.Windows.Forms.ToolStripButton btnOk;
        private System.Windows.Forms.ToolStripButton btnCancelar;
        private Timer pesquisaTimer;
    }
}
