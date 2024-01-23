
namespace BonifacioEntregas
{
    partial class operLancamento
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.panel1 = new System.Windows.Forms.Panel();
            this.dtpDataValidadeCNH = new System.Windows.Forms.DateTimePicker();
            this.txObs = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txCompra = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cmbCliente = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.lbMoto = new System.Windows.Forms.Label();
            this.cmbMotoBoy = new System.Windows.Forms.ComboBox();
            this.cmbFormaPagamento = new System.Windows.Forms.ComboBox();
            this.txtValor = new System.Windows.Forms.TextBox();
            this.btnLimpar = new System.Windows.Forms.Button();
            this.btnAdicionar = new System.Windows.Forms.Button();
            this.btnFiltrar = new System.Windows.Forms.Button();
            this.btnNovoCliente = new System.Windows.Forms.Button();
            this.dataGrid1 = new SourceGrid.DataGrid();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.dtpDataValidadeCNH);
            this.panel1.Controls.Add(this.txObs);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.txCompra);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.cmbCliente);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.lbMoto);
            this.panel1.Controls.Add(this.cmbMotoBoy);
            this.panel1.Controls.Add(this.cmbFormaPagamento);
            this.panel1.Controls.Add(this.txtValor);
            this.panel1.Controls.Add(this.btnLimpar);
            this.panel1.Controls.Add(this.btnAdicionar);
            this.panel1.Controls.Add(this.btnFiltrar);
            this.panel1.Controls.Add(this.btnNovoCliente);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(779, 101);
            this.panel1.TabIndex = 11;
            // 
            // dtpDataValidadeCNH
            // 
            this.dtpDataValidadeCNH.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.dtpDataValidadeCNH.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpDataValidadeCNH.Location = new System.Drawing.Point(603, 68);
            this.dtpDataValidadeCNH.Name = "dtpDataValidadeCNH";
            this.dtpDataValidadeCNH.Size = new System.Drawing.Size(83, 23);
            this.dtpDataValidadeCNH.TabIndex = 29;
            this.dtpDataValidadeCNH.Tag = "H";
            // 
            // txObs
            // 
            this.txObs.Location = new System.Drawing.Point(190, 71);
            this.txObs.Name = "txObs";
            this.txObs.Size = new System.Drawing.Size(395, 20);
            this.txObs.TabIndex = 28;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(101, 72);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(83, 16);
            this.label5.TabIndex = 27;
            this.label5.Text = "Observação";
            // 
            // txCompra
            // 
            this.txCompra.Location = new System.Drawing.Point(529, 38);
            this.txCompra.Name = "txCompra";
            this.txCompra.Size = new System.Drawing.Size(56, 20);
            this.txCompra.TabIndex = 26;
            this.txCompra.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txCompra.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtValor_KeyPress);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(467, 39);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(56, 16);
            this.label4.TabIndex = 25;
            this.label4.Text = "Compra";
            // 
            // cmbCliente
            // 
            this.cmbCliente.DisplayMember = "Nome";
            this.cmbCliente.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbCliente.Location = new System.Drawing.Point(82, 38);
            this.cmbCliente.Name = "cmbCliente";
            this.cmbCliente.Size = new System.Drawing.Size(266, 21);
            this.cmbCliente.TabIndex = 24;
            this.cmbCliente.ValueMember = "Id";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(30, 40);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(49, 16);
            this.label3.TabIndex = 23;
            this.label3.Text = "Cliente";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(483, 13);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(40, 16);
            this.label2.TabIndex = 22;
            this.label2.Text = "Valor";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(209, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(139, 16);
            this.label1.TabIndex = 21;
            this.label1.Text = "Forma de Pagamento";
            // 
            // lbMoto
            // 
            this.lbMoto.AutoSize = true;
            this.lbMoto.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbMoto.Location = new System.Drawing.Point(9, 12);
            this.lbMoto.Name = "lbMoto";
            this.lbMoto.Size = new System.Drawing.Size(61, 16);
            this.lbMoto.TabIndex = 20;
            this.lbMoto.Text = "Motoboy";
            // 
            // cmbMotoBoy
            // 
            this.cmbMotoBoy.DisplayMember = "Nome";
            this.cmbMotoBoy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbMotoBoy.Location = new System.Drawing.Point(82, 13);
            this.cmbMotoBoy.Name = "cmbMotoBoy";
            this.cmbMotoBoy.Size = new System.Drawing.Size(121, 21);
            this.cmbMotoBoy.TabIndex = 12;
            this.cmbMotoBoy.ValueMember = "Id";
            // 
            // cmbFormaPagamento
            // 
            this.cmbFormaPagamento.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbFormaPagamento.Items.AddRange(new object[] {
            "Anotado",
            "Cartão,",
            "Dinheiro",
            "Pix",
            "Troca"});
            this.cmbFormaPagamento.Location = new System.Drawing.Point(354, 11);
            this.cmbFormaPagamento.Name = "cmbFormaPagamento";
            this.cmbFormaPagamento.Size = new System.Drawing.Size(121, 21);
            this.cmbFormaPagamento.TabIndex = 13;
            // 
            // txtValor
            // 
            this.txtValor.Location = new System.Drawing.Point(529, 11);
            this.txtValor.Name = "txtValor";
            this.txtValor.Size = new System.Drawing.Size(56, 20);
            this.txtValor.TabIndex = 14;
            this.txtValor.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtValor_KeyPress);
            // 
            // btnLimpar
            // 
            this.btnLimpar.Location = new System.Drawing.Point(692, 10);
            this.btnLimpar.Name = "btnLimpar";
            this.btnLimpar.Size = new System.Drawing.Size(75, 23);
            this.btnLimpar.TabIndex = 16;
            this.btnLimpar.Text = "Limpar";
            // 
            // btnAdicionar
            // 
            this.btnAdicionar.Enabled = false;
            this.btnAdicionar.Location = new System.Drawing.Point(692, 40);
            this.btnAdicionar.Name = "btnAdicionar";
            this.btnAdicionar.Size = new System.Drawing.Size(75, 23);
            this.btnAdicionar.TabIndex = 17;
            this.btnAdicionar.Text = "Adicionar";
            this.btnAdicionar.Click += new System.EventHandler(this.btnAdicionar_Click);
            // 
            // btnFiltrar
            // 
            this.btnFiltrar.Location = new System.Drawing.Point(692, 69);
            this.btnFiltrar.Name = "btnFiltrar";
            this.btnFiltrar.Size = new System.Drawing.Size(75, 23);
            this.btnFiltrar.TabIndex = 18;
            this.btnFiltrar.Text = "Filtrar";
            // 
            // btnNovoCliente
            // 
            this.btnNovoCliente.Location = new System.Drawing.Point(12, 69);
            this.btnNovoCliente.Name = "btnNovoCliente";
            this.btnNovoCliente.Size = new System.Drawing.Size(83, 23);
            this.btnNovoCliente.TabIndex = 19;
            this.btnNovoCliente.Text = "Novo Cliente";
            // 
            // dataGrid1
            // 
            this.dataGrid1.DeleteQuestionMessage = "Are you sure to delete all the selected rows?";
            this.dataGrid1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGrid1.EnableSort = false;
            this.dataGrid1.FixedRows = 1;
            this.dataGrid1.Location = new System.Drawing.Point(0, 101);
            this.dataGrid1.Name = "dataGrid1";
            this.dataGrid1.SelectionMode = SourceGrid.GridSelectionMode.Row;
            this.dataGrid1.Size = new System.Drawing.Size(779, 460);
            this.dataGrid1.TabIndex = 12;
            this.dataGrid1.TabStop = true;
            this.dataGrid1.ToolTipText = "";
            // 
            // operLancamento
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(779, 561);
            this.Controls.Add(this.dataGrid1);
            this.Controls.Add(this.panel1);
            this.MaximizeBox = false;
            this.Name = "operLancamento";
            this.Text = "Lançamento de Entregas";
            this.Load += new System.EventHandler(this.operLancamento_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ComboBox cmbFormaPagamento;
        private System.Windows.Forms.TextBox txtValor;
        private System.Windows.Forms.Button btnLimpar;
        private System.Windows.Forms.Button btnAdicionar;
        private System.Windows.Forms.Button btnFiltrar;
        private System.Windows.Forms.Button btnNovoCliente;
        private SourceGrid.DataGrid dataGrid1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lbMoto;
        private System.Windows.Forms.TextBox txObs;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txCompra;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cmbCliente;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cmbMotoBoy;
        private System.Windows.Forms.DateTimePicker dtpDataValidadeCNH;
    }
}