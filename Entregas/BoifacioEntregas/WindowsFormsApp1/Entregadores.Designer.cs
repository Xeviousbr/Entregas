namespace BonifacioEntregas
{
    partial class Form2
    {
        private System.Windows.Forms.Label lblNome;
        private System.Windows.Forms.TextBox txtNome;
        private System.Windows.Forms.Label lblTelefone;
        private System.Windows.Forms.TextBox txtTelefone;
        private System.Windows.Forms.Label lblCNH;
        private System.Windows.Forms.TextBox txtCNH;
        private System.Windows.Forms.Label lblValidadeCNH;
        private System.Windows.Forms.DateTimePicker dtpDataValidadeCNH;


        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form2));
            this.cntrole1 = new ATCRecordNavigator.Cntrole();
            this.lblNome = new System.Windows.Forms.Label();
            this.txtNome = new System.Windows.Forms.TextBox();
            this.lblTelefone = new System.Windows.Forms.Label();
            this.txtTelefone = new System.Windows.Forms.TextBox();
            this.lblCNH = new System.Windows.Forms.Label();
            this.txtCNH = new System.Windows.Forms.TextBox();
            this.lblValidadeCNH = new System.Windows.Forms.Label();
            this.dtpDataValidadeCNH = new System.Windows.Forms.DateTimePicker();
            this.SuspendLayout();
            // 
            // cntrole1
            // 
            this.cntrole1.Dock = System.Windows.Forms.DockStyle.Top;
            this.cntrole1.EmAdicao = false;
            this.cntrole1.EmEdicao = false;
            this.cntrole1.Location = new System.Drawing.Point(0, 0);
            this.cntrole1.Name = "cntrole1";
            this.cntrole1.Primeiro = false;
            this.cntrole1.Size = new System.Drawing.Size(331, 54);
            this.cntrole1.TabIndex = 0;
            this.cntrole1.Ultimo = false;
            this.cntrole1.AcaoRealizada += new System.EventHandler<AcaoEventArgs>(this.cntrole1_AcaoRealizada);
            this.cntrole1.Load += new System.EventHandler(this.cntrole1_Load);
            // 
            // lblNome
            // 
            this.lblNome.AutoSize = true;
            this.lblNome.Location = new System.Drawing.Point(30, 70);
            this.lblNome.Name = "lblNome";
            this.lblNome.Size = new System.Drawing.Size(35, 13);
            this.lblNome.TabIndex = 0;
            this.lblNome.Text = "Nome";
            // 
            // txtNome
            // 
            this.txtNome.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtNome.Location = new System.Drawing.Point(30, 90);
            this.txtNome.MaxLength = 100;
            this.txtNome.Name = "txtNome";
            this.txtNome.Size = new System.Drawing.Size(250, 23);
            this.txtNome.TabIndex = 1;
            this.txtNome.Tag = "Nome";
            // 
            // lblTelefone
            // 
            this.lblTelefone.AutoSize = true;
            this.lblTelefone.Location = new System.Drawing.Point(30, 120);
            this.lblTelefone.Name = "lblTelefone";
            this.lblTelefone.Size = new System.Drawing.Size(49, 13);
            this.lblTelefone.TabIndex = 2;
            this.lblTelefone.Text = "Telefone";
            // 
            // txtTelefone
            // 
            this.txtTelefone.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtTelefone.Location = new System.Drawing.Point(30, 140);
            this.txtTelefone.MaxLength = 20;
            this.txtTelefone.Name = "txtTelefone";
            this.txtTelefone.Size = new System.Drawing.Size(108, 23);
            this.txtTelefone.TabIndex = 3;
            this.txtTelefone.Tag = "Telefone";
            // 
            // lblCNH
            // 
            this.lblCNH.AutoSize = true;
            this.lblCNH.Location = new System.Drawing.Point(30, 175);
            this.lblCNH.Name = "lblCNH";
            this.lblCNH.Size = new System.Drawing.Size(30, 13);
            this.lblCNH.TabIndex = 6;
            this.lblCNH.Text = "CNH";
            // 
            // txtCNH
            // 
            this.txtCNH.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtCNH.Location = new System.Drawing.Point(30, 195);
            this.txtCNH.MaxLength = 12;
            this.txtCNH.Name = "txtCNH";
            this.txtCNH.Size = new System.Drawing.Size(108, 23);
            this.txtCNH.TabIndex = 7;
            this.txtCNH.Tag = "CNH";
            // 
            // lblValidadeCNH
            // 
            this.lblValidadeCNH.AutoSize = true;
            this.lblValidadeCNH.Location = new System.Drawing.Point(206, 175);
            this.lblValidadeCNH.Name = "lblValidadeCNH";
            this.lblValidadeCNH.Size = new System.Drawing.Size(74, 13);
            this.lblValidadeCNH.TabIndex = 8;
            this.lblValidadeCNH.Text = "Validade CNH";
            // 
            // dtpDataValidadeCNH
            // 
            this.dtpDataValidadeCNH.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.dtpDataValidadeCNH.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpDataValidadeCNH.Location = new System.Drawing.Point(206, 195);
            this.dtpDataValidadeCNH.Name = "dtpDataValidadeCNH";
            this.dtpDataValidadeCNH.Size = new System.Drawing.Size(100, 23);
            this.dtpDataValidadeCNH.TabIndex = 9;
            this.dtpDataValidadeCNH.Tag = "A";
            this.dtpDataValidadeCNH.ValueChanged += new System.EventHandler(this.dtpValidadeCNH_ValueChanged);
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(331, 237);
            this.Controls.Add(this.cntrole1);
            this.Controls.Add(this.lblNome);
            this.Controls.Add(this.txtNome);
            this.Controls.Add(this.lblTelefone);
            this.Controls.Add(this.txtTelefone);
            this.Controls.Add(this.lblCNH);
            this.Controls.Add(this.txtCNH);
            this.Controls.Add(this.lblValidadeCNH);
            this.Controls.Add(this.dtpDataValidadeCNH);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.Name = "Form2";
            this.Text = "Cadastro de Entregador";
            this.KeyUp += new System.Windows.Forms.KeyEventHandler(this.Teclou);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private ATCRecordNavigator.Cntrole cntrole1;
    }
}
