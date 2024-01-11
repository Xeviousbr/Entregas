
using System;
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

            // Configura o campo de texto
            this.txtPesquisar.Size = new System.Drawing.Size(100, 25); // ajuste o tamanho conforme necessário
            this.txtPesquisar.Visible = false; // inicialmente oculto

            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.btnPesquisa = new System.Windows.Forms.ToolStripButton();
            this.btnParaTras = new System.Windows.Forms.ToolStripButton();
            this.btnParaFrente = new System.Windows.Forms.ToolStripButton();
            this.btnEditar = new System.Windows.Forms.ToolStripButton();
            this.btnApagar = new System.Windows.Forms.ToolStripButton();

            this.btnPesquisarAdicional.Text = "Pesquisar";
            this.btnPesquisarAdicional.Click += new System.EventHandler(this.btnPesquisarAdicional_Click);
            this.btnPesquisarAdicional.Visible = false; // inicialmente oculto

            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStrip1
            // 
            this.toolStrip1.ImageScalingSize = new System.Drawing.Size(48, 48);
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnPesquisa,
            this.btnParaFrente,
            this.btnParaTras,
            this.btnEditar,
            this.btnApagar,
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
            this.btnPesquisa.Click += new System.EventHandler(this.btnPesquisa_Click);
            // 
            // btnParaTras
            // 
            this.btnParaTras.Image = ((System.Drawing.Image)(resources.GetObject("btnParaTras.Image")));
            this.btnParaTras.Name = "btnParaTras";
            this.btnParaTras.Size = new System.Drawing.Size(52, 52);
            // 
            // btnParaFrente
            // 
            this.btnParaFrente.Image = ((System.Drawing.Image)(resources.GetObject("btnParaFrente.Image")));
            this.btnParaFrente.Name = "btnParaFrente";
            this.btnParaFrente.Size = new System.Drawing.Size(52, 52);
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

        private void btnPesquisarAdicional_Click(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        private void btnPesquisa_Click(object sender, EventArgs e)
        {
            // Alternar visibilidade dos botões originais
            this.btnParaTras.Visible = !this.btnParaTras.Visible;
            this.btnParaFrente.Visible = !this.btnParaFrente.Visible;
            this.btnEditar.Visible = !this.btnEditar.Visible;
            this.btnApagar.Visible = !this.btnApagar.Visible;

            // Alternar visibilidade do campo de texto e do botão de pesquisa adicional
            this.txtPesquisar.Visible = !this.txtPesquisar.Visible;
            this.btnPesquisarAdicional.Visible = !this.btnPesquisarAdicional.Visible;

            // Se o campo de texto está visível, foca nele
            if (this.txtPesquisar.Visible)
            {
                this.txtPesquisar.Focus();
            }
        }
        #endregion

        private System.Windows.Forms.ToolTip toolTip1;
        private ToolStrip toolStrip1;
        private ToolStripButton btnPesquisa;        
        private ToolStripButton btnParaFrente;
        private ToolStripButton btnParaTras;
        private ToolStripButton btnEditar;
        private ToolStripButton btnApagar;
        private ToolStripTextBox txtPesquisar;
        private ToolStripButton btnPesquisarAdicional;
    }
}
