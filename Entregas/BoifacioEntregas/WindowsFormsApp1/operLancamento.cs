using BonifacioEntregas.dao;
using System;
using System.Collections.Generic;
using System.Data;
using System.Reflection;
using System.Windows.Forms;

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
            dataGrid1.Columns[1].Width = 70;
            dataGrid1.Columns[2].Width = 110;

            dataGrid1.Columns[3].Width = 50;
            //DataGridColumn colunaValor = new DataGridColumn(dataGrid1, null, new SourceGrid.Cells.DataGrid.Cell(), "Valor");
            //colunaValor.DataCell.AddController(new SourceGrid.Cells.Controllers.Unselectable());
            //colunaValor.DataCell.View = new SourceGrid.Cells.Views.Cell();
            //colunaValor.DataCell.View.TextAlignment = DevAge.Drawing.ContentAlignment.MiddleRight;
            //colunaValor.DataCell.View.Font = new Font("Arial", 10);
            //colunaValor.DataCell.View.ForeColor = Color.Black;
            //colunaValor.DataCell.View.BackColor = Color.White;
            //colunaValor.DataCell.View.Border = DevAge.Drawing.RectangleBorder.NoBorder;
            //// colunaValor.DataCell.View..Padding = new DevAge.Drawing.Padding(5);
            //// colunaValor.DataCell.View.
            ////.Format = "C2"; // Formato de moeda com duas casas decimais
            //dataGrid1.Columns.Insert(4, colunaValor);
            //    //.Add(colunaValor);

            dataGrid1.Columns[4].Width = 90;
            dataGrid1.Columns[5].Width = 70;
            dataGrid1.Columns[6].Width = 310;

            // dataGrid1.Columns[3].PropertyColumn.
            //.DefaultCellStyle.Format = "C2";
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
            List<T> lista = ConvertDataTableToList<T>(dados);
            comboBox.DataSource = lista;
        }

        //public List<T> ConvertDataTableToList<T>(DataTable dataTable) where T : new()
        //{
        //    List<T> list = new List<T>();
        //    foreach (DataRow row in dataTable.Rows)
        //    {
        //        T item = new T();
        //        foreach (DataColumn column in dataTable.Columns)
        //        {
        //            string propertyName = column.ColumnName;
        //            PropertyInfo property = typeof(T).GetProperty(propertyName);
        //            if (property != null && row[column] != DBNull.Value)
        //            {
        //                Type propertyType = property.PropertyType;
        //                object value = Convert.ChangeType(row[column], propertyType);
        //                property.SetValue(item, value, null);
        //            }
        //        }
        //        list.Add(item);
        //    }
        //    return list;
        //}

        public List<T> ConvertDataTableToList<T>(DataTable dataTable) where T : new()
        {
            List<T> list = new List<T>();
            foreach (DataRow row in dataTable.Rows)
            {
                T item = new T();
                foreach (DataColumn column in dataTable.Columns)
                {
                    string propertyName = column.ColumnName;
                    PropertyInfo property = typeof(T).GetProperty(propertyName);
                    if (property != null && row[column] != DBNull.Value)
                    {
                        property.SetValue(item, row[column], null);
                    }
                }
                list.Add(item);
            }
            return list;
        }

        private void btnAdicionar_Click(object sender, EventArgs e)
        {
            int idBoy = Convert.ToInt32(cmbMotoBoy.SelectedValue);
            int idForma = Convert.ToInt32(cmbFormaPagamento.SelectedValue);
            int idCliente = Convert.ToInt32(cmbCliente.SelectedValue);
            float valor;
            if (!float.TryParse(txtValor.Text, out valor))
            {
                valor = 0; // ou manipule o erro conforme necessário
            }

            float compra;
            if (!float.TryParse(txCompra.Text, out compra))
            {
                compra = 0; // ou manipule o erro conforme necessário
            }
            string obs = txObs.Text;
            entregasDAO.Adiciona(idBoy, idForma, valor, idCliente, compra, obs);
        }
    }
}
