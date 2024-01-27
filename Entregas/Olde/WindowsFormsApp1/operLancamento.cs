 using BonifacioEntregas.dao;
using System;
using System.Collections.Generic;
using System.Data;
using System.Reflection;
using System.Windows.Forms;
using TeleBonifacio.tb;

namespace BonifacioEntregas
{
    public partial class operLancamento : Form
    {
        private EntregasDAO entregasDAO;
        private int iID = 0;

        public operLancamento()
        {
            InitializeComponent();
        }

        private void txtValor_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != ',' && e.KeyChar != '.')
            {
                e.Handled = true;
            }
        }

        private void operLancamento_Load(object sender, EventArgs e)
        {
            EntregadorDAO Entregador = new EntregadorDAO();
            ClienteDAO Cliente = new ClienteDAO();
            CarregarComboBox<tb.Entregador>(cmbMotoBoy, Entregador);
            CarregarComboBox<tb.Cliente>(cmbCliente, Cliente);
            CarregaGrid(null);
            ConfigurarGrid();
        }
        private void ConfigurarGrid()
        {
            dataGrid1.Columns[0].Width = 0;
            dataGrid1.Columns[1].Width = 75;
            dataGrid1.Columns[2].Width = 110;
            dataGrid1.Columns[3].Width = 50;
            dataGrid1.Columns[4].Width = 50;    // Desconto
            dataGrid1.Columns[5].Width = 90;
            dataGrid1.Columns[6].Width = 70;
            dataGrid1.Columns[7].Width = 290;   // Cliente
            dataGrid1.Columns[8].Width = 90;    // Obs
            dataGrid1.Columns[9].Width = 0;
            dataGrid1.Columns[10].Width = 0;
            dataGrid1.Columns[11].Width = 0;
            dataGrid1.Invalidate();
        }

        private void CarregaGrid(DateTime? DT)
        {
            entregasDAO = new EntregasDAO();            
            DataTable dados = entregasDAO.getDados(DT);
            DevAge.ComponentModel.BoundDataView boundDataView = new DevAge.ComponentModel.BoundDataView(dados.DefaultView);
            dataGrid1.DataSource = boundDataView;
        }

        private void CarregarComboBox<T>(ComboBox comboBox, BaseDAO classe) where T : tb.IDataEntity, new()
        {
            DataTable dados = classe.getDadosOrdenados();
            List<ComboBoxItem> lista = new List<ComboBoxItem>();
            foreach (DataRow row in dados.Rows)
            {
                int id = Convert.ToInt32(row["id"]);
                string nome = row["Nome"].ToString();

                ComboBoxItem item = new ComboBoxItem(id, nome);
                lista.Add(item);
            }
            comboBox.DataSource = lista;
            comboBox.DisplayMember = "Nome";
            comboBox.ValueMember = "Id";
        }

        private void btnAdicionar_Click(object sender, EventArgs e)
        {
            int idBoy = Convert.ToInt32(cmbMotoBoy.SelectedValue);
            int idForma = Convert.ToInt32(cmbFormaPagamento.SelectedIndex);
            int idCliente = Convert.ToInt32(cmbCliente.SelectedValue);
            float valor;
            if (!float.TryParse(txtValor.Text, out valor))
            {
                valor = 0;
            }
            float compra;
            if (!float.TryParse(txCompra.Text, out compra))
            {
                compra = 0;
            }
            string obs = txObs.Text;
            float desc;            
            if (!float.TryParse(txDesc.Text, out desc))
            {
                desc = 0;
            }                        
            if (btnAdicionar.Text == "Salvar")
            {
                entregasDAO.Edita(this.iID, idBoy, idForma, valor, idCliente, compra, obs, desc);
                btnAdicionar.Text = "Adicionar";
            } else
            {
                entregasDAO.Adiciona(idBoy, idForma, valor, idCliente, compra, obs, desc);
            }
            CarregaGrid(null);
//             Limpar();
        }

        #region Criticas
        private void VeSeHab()
        {
            bool OK = true;
            if (cmbFormaPagamento.SelectedIndex == -1)
            {
                OK = false;
            }
            if (txtValor.Text == "")
            {
                OK = false;
            }
            btnAdicionar.Enabled = OK;
        }

        #endregion

        private void cmbMotoBoy_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                string searchText = cmbMotoBoy.Text.Trim();
                cmbMotoBoy.SelectedValue = int.Parse(searchText);
            }
        }

        private void cmbCliente_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                dao.ClienteDAO cCli = new dao.ClienteDAO();
                string searchText = cmbCliente.Text.Trim();
                int idCli = cCli.RetIdNrAlter(searchText);
                if (idCli>0)
                {
                    cmbCliente.SelectedValue = idCli;
                }                
            }
        }

        private void txtValor_KeyUp(object sender, KeyEventArgs e)
        {
            float valor = gen.LeValor(txtValor.Text);
            float compra = gen.LeValor(txCompra.Text);
            float desc = gen.LeValor(txDesc.Text);
            float total = valor + compra - desc;
            if (total > 0)
            {
                lbTotal.Text = total.ToString("C");
            } else
            {
                lbTotal.Text = "";
            }
            VeSeHab();
        }

        private void Limpar()
        {
            cmbMotoBoy.SelectedIndex = -1;
            cmbCliente.SelectedIndex = -1;
            cmbFormaPagamento.SelectedIndex = -1;
            txtValor.Text = "";
            txCompra.Text = "";
            txDesc.Text = "";
        }

        private void btnLimpar_Click(object sender, EventArgs e)
        {
            Limpar();
        }

        private void dataGrid1_Click(object sender, EventArgs e)
        {
            SourceGrid.DataGrid grid = (SourceGrid.DataGrid)sender;
            if (grid != null && grid.Rows.Count > 0)
            {
                SourceGrid.Position position = grid.Selection.ActivePosition;
                if (position != SourceGrid.Position.Empty)
                {
                    this.iID = gen.ConvOjbInt(((DataRowView)grid.SelectedDataRows[0]).Row.ItemArray[0]);
                    txtValor.Text = gen.ConvOjbStr(((DataRowView)grid.SelectedDataRows[0]).Row.ItemArray[3]);

                    txDesc.Text = gen.ConvOjbStr(((DataRowView)grid.SelectedDataRows[0]).Row.ItemArray[4]);

                    txCompra.Text = gen.ConvOjbStr(((DataRowView)grid.SelectedDataRows[0]).Row.ItemArray[6]);                    
                    txObs.Text = gen.ConvOjbStr(((DataRowView)grid.SelectedDataRows[0]).Row.ItemArray[8]);
                    cmbMotoBoy.SelectedValue = gen.ConvOjbInt(((DataRowView)grid.SelectedDataRows[0]).Row.ItemArray[9]);
                    cmbCliente.SelectedValue = gen.ConvOjbInt(((DataRowView)grid.SelectedDataRows[0]).Row.ItemArray[10]);
                    cmbFormaPagamento.SelectedIndex = gen.ConvOjbInt(((DataRowView)grid.SelectedDataRows[0]).Row.ItemArray[11]);
                    btnAdicionar.Text = "Salvar";
                }
            }
        }

        private void cmbFormaPagamento_SelectedIndexChanged(object sender, EventArgs e)
        {
            VeSeHab();
        }

        private void btnFiltrar_Click(object sender, EventArgs e)
        {
            DateTime DT = dtpData.Value;
            CarregaGrid(DT);
        }

        private void operLancamento_Resize(object sender, EventArgs e)
        {
            dataGrid1.Width = this.Width;
        }
    }
}

