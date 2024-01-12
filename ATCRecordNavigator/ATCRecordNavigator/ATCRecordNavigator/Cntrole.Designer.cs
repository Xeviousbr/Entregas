
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Cntrole));
            this.txtPesquisar = new System.Windows.Forms.ToolStripTextBox();
            this.btnPesquisarAdicional = new System.Windows.Forms.ToolStripButton();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.btnPesquisa = new System.Windows.Forms.ToolStripButton();
            this.btnAdicionar = new System.Windows.Forms.ToolStripButton();
            this.btnEditar = new System.Windows.Forms.ToolStripButton();
            this.btnApagar = new System.Windows.Forms.ToolStripButton();
            this.btnParaFrente = new System.Windows.Forms.ToolStripButton();
            this.btnParaTras = new System.Windows.Forms.ToolStripButton();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtPesquisar
            // 
            this.txtPesquisar.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtPesquisar.Name = "txtPesquisar";
            this.txtPesquisar.Size = new System.Drawing.Size(100, 55);
            this.txtPesquisar.Visible = false;
            // 
            // btnPesquisarAdicional
            // 
            this.btnPesquisarAdicional.Name = "btnPesquisarAdicional";
            this.btnPesquisarAdicional.Size = new System.Drawing.Size(61, 52);
            this.btnPesquisarAdicional.Text = "Pesquisar";
            this.btnPesquisarAdicional.Visible = false;
            this.btnPesquisarAdicional.Click += new System.EventHandler(this.btnPesquisarAdicional_Click);
            // 
            // toolStrip1
            // 
            this.toolStrip1.ImageScalingSize = new System.Drawing.Size(48, 48);
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnPesquisa,
            this.btnAdicionar,
            this.btnEditar,
            this.btnApagar,
            this.btnParaTras,
            this.btnParaFrente,
            this.txtPesquisar,
            this.btnPesquisarAdicional});
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
            // btnEditar
            // 
            this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
            this.btnEditar.Name = "btnEditar";
            this.btnEditar.Size = new System.Drawing.Size(52, 52);
            // 
            // btnApagar
            // 
            this.btnApagar.Image = ((System.Drawing.Image)(resources.GetObject("btnApagar.Image")));
            this.btnApagar.Name = "btnApagar";
            this.btnApagar.Size = new System.Drawing.Size(52, 52);
            // 
            // btnParaFrente
            // 
            this.btnParaFrente.Image = ((System.Drawing.Image)(resources.GetObject("btnParaFrente.Image")));
            this.btnParaFrente.Name = "btnParaFrente";
            this.btnParaFrente.Size = new System.Drawing.Size(52, 52);
            // 
            // btnParaTras
            // 
            this.btnParaTras.Image = ((System.Drawing.Image)(resources.GetObject("btnParaTras.Image")));
            this.btnParaTras.Name = "btnParaTras";
            this.btnParaTras.Size = new System.Drawing.Size(52, 52);
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
            throw new NotImplementedException();
        }

        private void btnPesquisarAdicional_Click(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        private void btnPesquisa_Click(object sender, EventArgs e)
        {
            // Verifica o estado atual do campo de texto para determinar a ação
            bool isSearchVisible = this.txtPesquisar.Visible;

            // Esconde ou mostra os botões padrão
            this.btnParaFrente.Visible = isSearchVisible;
            this.btnParaTras.Visible = isSearchVisible;
            this.btnEditar.Visible = isSearchVisible;
            this.btnApagar.Visible = isSearchVisible;

            // Alterna a visibilidade do campo de texto e do botão de pesquisa adicional
            this.txtPesquisar.Visible = !isSearchVisible;
            this.btnPesquisarAdicional.Visible = !isSearchVisible;

            if (this.txtPesquisar.Visible)
            {
                // Configurações quando o campo de pesquisa está visível
                this.txtPesquisar.AutoSize = false;
                int espacoParaBotao = 100; // Ajuste conforme necessário
                this.txtPesquisar.Width = this.toolStrip1.Width - espacoParaBotao;
                this.txtPesquisar.Font = new Font(this.txtPesquisar.Font.FontFamily, 12);

                this.btnPesquisarAdicional.Alignment = ToolStripItemAlignment.Right;
                this.btnPesquisarAdicional.Font = new Font(this.btnPesquisarAdicional.Font.FontFamily, 12);

                this.txtPesquisar.Focus();
            }
            else
            {
                // Reverter as configurações quando a pesquisa não está visível
                this.txtPesquisar.AutoSize = true;
                this.txtPesquisar.Font = new Font(this.txtPesquisar.Font.FontFamily, 8.25F);
                this.btnPesquisarAdicional.Alignment = ToolStripItemAlignment.Left;
                this.btnPesquisarAdicional.Font = new Font(this.btnPesquisarAdicional.Font.FontFamily, 8.25F);
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
        private ToolStripButton btnPesquisarAdicional;
        private System.Windows.Forms.ToolStripButton btnAdicionar;

    }
}
