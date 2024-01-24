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
            CarregaGrid();
            ConfigurarGrid();
        }
        private void ConfigurarGrid()
        {
            dataGrid1.Columns[0].Width = 0;
            dataGrid1.Columns[1].Width = 75;
            dataGrid1.Columns[2].Width = 110;
            dataGrid1.Columns[3].Width = 50;
            dataGrid1.Columns[4].Width = 90;
            dataGrid1.Columns[5].Width = 70;
            dataGrid1.Columns[6].Width = 310;
            dataGrid1.Invalidate();
        }

        private void CarregaGrid()
        {
            entregasDAO = new EntregasDAO();
            DataTable dados = entregasDAO.getDados();
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
            entregasDAO.Adiciona(idBoy, idForma, valor, idCliente, compra, obs);
            CarregaGrid();
        }

        #region Criticas

        private void VeSeHab()
        {
            // btnAdicionar
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
                string searchText = cmbCliente.Text.Trim();
                cmbCliente.SelectedValue = int.Parse(searchText);
            }
        }
    }
}

