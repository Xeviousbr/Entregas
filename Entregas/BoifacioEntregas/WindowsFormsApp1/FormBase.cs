﻿using BonifacioEntregas.tb;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SourceGrid;
using System.Diagnostics;

namespace BonifacioEntregas
{
    public partial class FormBase : Form
    {

        protected int Direcao = 0;
        protected bool EmAdicao = false;
        protected bool Mostrando = false;
        protected bool Pesquisando = false;
        protected dao.BaseDAO DAO;
        protected tb.IDataEntity reg;
        private List<CampoTagInfo> tagsDosCampos;
        private int lastColumnClick = -1;
        private DateTime lastClickTime = DateTime.MinValue;
        private System.Windows.Forms.DataGrid dataGrid;
        private bool GridCarregada = false;

        public FormBase()
        {
            InitializeComponent();
            tagsDosCampos = new List<tb.CampoTagInfo>();
        }

        private void InitializeDataGrid()
        {
            DataTable dataTable = DAO.CarregarDados();
        }

        #region Campos

        protected bool Mostra()
        {
            if (reg == null)
            {
                return false;
            }
            else
            {
                Pesquisando = false;
                Mostrando = true;
                foreach (Control control in this.Controls)
                {
                    if (control is TextBox textBox)
                    {
                        ProcessarTextBox(textBox);
                    }
                    else if (control is DateTimePicker dateTimePicker)
                    {
                        ProcessarDateTimePicker(dateTimePicker);
                    }
                }
                Mostrando = false;
                return true;
            }
        }

        protected void ProcessarTextBox(TextBox textBox)
        {
            string propertyName = textBox.Name.Substring(3); // Remove o prefixo 'txt'
            PropertyInfo propertyInfo = reg.GetType().GetProperty(propertyName);
            if (propertyInfo != null)
            {
                string valor = propertyInfo.GetValue(reg, null)?.ToString() ?? string.Empty;
                textBox.Text = valor;
            }
        }

        private void ProcessarDateTimePicker(DateTimePicker dtpControl)
        {
            string propertyName = dtpControl.Name.Substring(3);
            PropertyInfo propertyInfo = reg.GetType().GetProperty(propertyName);
            if (propertyInfo != null && propertyInfo.PropertyType == typeof(DateTime) || propertyInfo.PropertyType == typeof(DateTime?))
            {
                DateTime? data = propertyInfo.GetValue(reg, null) as DateTime?;
                if (!data.HasValue || data.Value == DateTime.MinValue)
                {
                    dtpControl.CustomFormat = " ";
                    dtpControl.Format = DateTimePickerFormat.Custom;
                }
                else
                {
                    dtpControl.Value = data.Value;
                    dtpControl.Format = DateTimePickerFormat.Short;
                }
            }
        }

        private void MapearCamposParaModelo(dao.BaseDAO reg)
        {
            foreach (Control control in this.Controls)
            {
                try
                {
                    if (control is TextBox textBox)
                    {
                        MapearTextBoxParaModelo(textBox, reg);
                    }
                    else if (control is DateTimePicker dateTimePicker)
                    {
                        MapearDateTimePickerParaModelo(dateTimePicker, reg);
                    }
                }
                catch (Exception ex)
                {
                    string x = ex.ToString();
                    // Tratamento adequado de exceções
                }
            }
        }

        private void MapearTextBoxParaModelo(TextBox textBox, dao.BaseDAO reg)
        {
            string propertyName = textBox.Name.Substring(3); // Remove o prefixo 'txt'
            PropertyInfo propertyInfo = reg.GetType().GetProperty(propertyName);
            if (propertyInfo == null)
            {
                propertyInfo.SetValue(reg, null, null);
            }
            else
            {
                propertyInfo.SetValue(reg, textBox.Text, null);
            }
        }

        private void MapearDateTimePickerParaModelo(DateTimePicker dtpControl, dao.BaseDAO reg)
        {
            string propertyName = dtpControl.Name.Substring(3); // Remove o prefixo 'dtp'
            PropertyInfo propertyInfo = reg.GetType().GetProperty(propertyName);
            if (propertyInfo != null && (propertyInfo.PropertyType == typeof(DateTime) || propertyInfo.PropertyType == typeof(DateTime?)))
            {
                if (dtpControl.Format != DateTimePickerFormat.Custom)
                {
                    propertyInfo.SetValue(reg, dtpControl.Value, null);
                }
                else
                {
                    propertyInfo.SetValue(reg, null, null); // ou DateTime.MinValue
                }
            }
        }

        #endregion

        #region TratamentoDeTela
        protected void LimparCampos()
        {
            foreach (Control control in this.Controls)
            {
                if (control is TextBox)
                {
                    control.Text = string.Empty;
                }
            }
        }

        public void ResetarAparenciaControles()
        {
            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is TextBox)
                {
                    ctrl.BackColor = SystemColors.Window; // Cor normal
                }
            }
        }

        protected void cntrole1_AcaoRealizada(object sender, AcaoEventArgs e, tb.IDataEntity entidade)
        {
            switch (e.Acao)
            {
                case "Adicionar":
                    LimparCampos();
                    EmAdicao = true;
                    break;
                case "Delete":
                    reg = DAO.Apagar(Direcao, entidade);
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
                    reg = DAO.ParaTraz();
                    if (!Mostra())
                    {
                        cntrole1.Ultimo = true;
                    }
                    break;
                case "ParaFrente":
                    Direcao = 1; ;
                    reg = DAO.ParaFrente();
                    if (!Mostra())
                    {
                        cntrole1.Primeiro = true;
                    }
                    break;
                case "Editar":
                    // this.Text = "clicou";
                    break;
                case "CANC":
                    Cancela();
                    break;
                case "OK":
                    Grava();
                    break;
                case "PesqON":
                    LigaGrid();
                    break;
                case "PesqAcionar":
                    PesqAcionar();
                    break;
                case "PesqOFF":
                    PesqOFF();
                    break;
                case "Pesquisar":
                    Pesquisar();
                    break;
            }
        }

        private void PesqOFF()
        {
            AlterarVisibilidadeControles(true);
            dataGrid.Visible = false;            
            Pesquisando = false;
        }

        private void Pesquisar()
        {
            string Pesquisar = cntrole1.Pesquisa;
            if (Pesquisar.Length > 1)
            {
                System.Data.DataTable Dados = DAO.Fitrar(Pesquisar);
                dataGrid.DataSource = Dados;
            }
        }

        private System.Data.DataTable Fitrar(string pesquisar)
        {
            throw new NotImplementedException();
        }

        #endregion

        #region Crud
        protected void Grava()
        {
            MapearCamposParaModelo(DAO);
            List<string> criticas = FazerCriticas(DAO);
            if (criticas.Count == 0)
            {
                DAO.Adicao = EmAdicao;
                DAO.Grava(DAO);
                EmAdicao = false;
                cntrole1.ModoNormal();
            }
            else
            {
                string mensagemCritica = string.Join("\n", criticas);
                MessageBox.Show(mensagemCritica, "Críticas", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        #endregion

        #region Criticas

        public List<string> FazerCriticas<T>(T objeto) where T : class
        {
            List<string> criticas = new List<string>();
            PropertyInfo[] propriedades = objeto.GetType().GetProperties();
            foreach (CampoTagInfo campoTag in tagsDosCampos)
            {
                PropertyInfo propriedade = propriedades.FirstOrDefault(p => p.Name.Equals(campoTag.Nome, StringComparison.OrdinalIgnoreCase));
                if (propriedade != null)
                {
                    if (campoTag.Tag == "O")
                    {
                        object valor = propriedade.GetValue(objeto);
                        if (valor == null || (valor is string && string.IsNullOrEmpty((string)valor)))
                        {
                            criticas.Add($"O campo {propriedade.Name} é obrigatório.");
                        }
                    }
                    else if (campoTag.Tag == "H" && propriedade.PropertyType == typeof(DateTime))
                    {
                        DateTime dataValor = (DateTime)propriedade.GetValue(objeto);
                        if (dataValor > DateTime.Today)
                        {
                            criticas.Add($"A data no campo {propriedade.Name} não pode ser posterior a hoje.");
                        }
                    }
                }
            }
            MarcarControlesComErro(criticas);
            return criticas;
        }

        public void LerTagsDosCamposDeTexto()
        {
            tagsDosCampos.Clear();
            foreach (Control control in Controls)
            {
                if (control is TextBox textBox)
                {
                    if (textBox.Tag != null)
                    {
                        tagsDosCampos.Add(new CampoTagInfo
                        {
                            Nome = textBox.Name.Substring(3),
                            Tag = textBox.Tag.ToString()
                        });
                    }
                }
                else if (control is DateTimePicker dtp)
                {
                    if (dtp.Tag != null)
                    {
                        tagsDosCampos.Add(new CampoTagInfo
                        {
                            Nome = dtp.Name.Substring(3),
                            Tag = dtp.Tag.ToString()
                        });
                    }
                }
            }
        }

        private void MarcarControlesComErro(List<string> criticas)
        {
            HashSet<string> camposComErro = new HashSet<string>();
            foreach (string critica in criticas)
            {
                string[] palavras = critica.Split(' ');
                int indiceCampo = Array.IndexOf(palavras, "campo");
                if (indiceCampo != -1 && indiceCampo + 1 < palavras.Length)
                {
                    string nomeCampo = palavras[indiceCampo + 1];
                    camposComErro.Add(nomeCampo);
                }
            }
            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is TextBox textBox)
                {
                    string nomeCampo = textBox.Name.Substring(3);
                    if (camposComErro.Contains(nomeCampo))
                    {
                        textBox.BackColor = Color.LightPink;
                    }
                    else
                    {
                        textBox.BackColor = SystemColors.Window;
                    }
                }
                else if (ctrl is DateTimePicker dtp)
                {
                    string nomeCampo = dtp.Name.Substring(3);
                    if (camposComErro.Contains(nomeCampo))
                    {
                        dtp.Font = new Font(dtp.Font.FontFamily, dtp.Font.Size, FontStyle.Bold);
                    }
                    else
                    {
                        dtp.Font = new Font(dtp.Font.FontFamily, dtp.Font.Size, FontStyle.Regular);
                    }
                }
            }
        }

        #endregion

        #region Pesquisa

        public void NrLinhas(int v)
        {
            DAO.SetarLinhas(v);
        }

        private System.Data.DataTable getDados()
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            System.Data.DataTable X = DAO.getDados();
            stopwatch.Stop();
            TimeSpan tempoDecorrido = stopwatch.Elapsed;
            string tempoStr = tempoDecorrido.ToString(@"hh\:mm\:ss\.fff");
            INI MeuIni = new INI();
            MeuIni.WriteString("Clientes", "Quantidade", X.Rows.Count.ToString());
            MeuIni.WriteString("Clientes", "Tempo Pesquisa no Cadastro", tempoStr);
            return X;
        }
        private void PesqAcionar()
        {
            AlterarVisibilidadeControles(true);
        }
        public void LigaGrid()
        {
            Pesquisando = true;
            AlterarVisibilidadeControles(false);
            if (!GridCarregada)
            {
                CriaGrid();
                GridCarregada = true;
            }
        }

        private void CriaGrid()
        {
            System.Data.DataTable Dados = getDados();
            dataGrid = new System.Windows.Forms.DataGrid();
            this.Controls.Add(dataGrid);
            dataGrid.DataSource = Dados;
            dataGrid.Name = "GRID";
            dataGrid.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom;
            dataGrid.ReadOnly = true;
            dataGrid.DoubleClick += new EventHandler(dataGrid_DoubleClick);
            int posY = cntrole1.Location.Y + cntrole1.Height;
            posY -= 20;
            int alturaDataGrid = this.ClientSize.Height - posY;
            dataGrid.SetBounds(0, posY, this.ClientSize.Width, alturaDataGrid);
            dataGrid.ColumnHeadersVisible = false;
        }

        private void dataGrid_DoubleClick(object sender, EventArgs e)
        {
            System.Windows.Forms.DataGrid grid = (System.Windows.Forms.DataGrid)sender;
            if (grid.CurrentRowIndex >= 0)
            {
                int rowIndex = grid.CurrentRowIndex;
                System.Data.DataRowView selectedRowView = (System.Data.DataRowView)grid.BindingContext[grid.DataSource].Current;
                object idValue = selectedRowView.Row["id"];
                CarregaRegistro(idValue.ToString());
            }
        }

        private void CarregaRegistro(string v)
        {
            reg = DAO.GetPeloID(v);
            Mostra();
            PesqAcionar();
            cntrole1.ControlesNormais();
        }

        protected void Cancela()
        {
            reg = DAO.GetEsse();
            ResetarAparenciaControles();
            Mostra();
        }

        private void AlterarVisibilidadeControles(bool visivel)
        {
            foreach (Control control in this.Controls)
            {
                switch (control.Name)
                {
                    case "cntrole1":
                        break;
                    case "GRID":
                        control.Visible = !visivel;
                        break;
                    default:
                        control.Visible = visivel;
                        break;
                }
            }
        }

        #endregion

    }
}
