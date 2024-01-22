using BonifacioEntregas.dao;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BonifacioEntregas
{
    public partial class operLancamento : Form
    {
        public operLancamento()
        {
            InitializeComponent();
        }

        private void txtValor_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != ',' && e.KeyChar != '.')
            {
                // Impede a inserção do caractere
                e.Handled = true;
            }
        }

        private void operLancamento_Load(object sender, EventArgs e)
        {
            
            EntregadorDAO Entregador = new EntregadorDAO();
            ClienteDAO Cliente = new ClienteDAO();
            DataTable dadosEntrega = Entregador.getDadosOrdenados();            
            List<tb.Entregador> listaMotoBoys = ConvertDataTableToList<tb.Entregador>(dadosEntrega);            
            this.cmbMotoBoy.DataSource = listaMotoBoys;
            this.cmbMotoBoy.DisplayMember = "Nome";
            this.cmbMotoBoy.ValueMember = "Id";            
            this.cmbCliente.DisplayMember = "Nome";
            this.cmbCliente.ValueMember = "Id";

            Stopwatch stopwatch = new Stopwatch();
            INI MeuIni = new INI();
            stopwatch.Start();
            DataTable dadosCliente = Cliente.getDadosOrdenados();
            MeuIni.WriteString("Clientes", "Quantidade", dadosCliente.Rows.Count.ToString());

            TimeSpan tempoDecorrido = stopwatch.Elapsed;
            string tempoStr = tempoDecorrido.ToString(@"hh\:mm\:ss\.fff");
            MeuIni.WriteString("Clientes", "getDadosOrdenados", tempoStr);

            List<tb.Cliente> listaClientes = ConvertDataTableToList<tb.Cliente>(dadosCliente);
            TimeSpan tempoConvert = stopwatch.Elapsed;
            string tempoConver = tempoConvert.ToString(@"hh\:mm\:ss\.fff");
            MeuIni.WriteString("Clientes", "ConvertDataTableToList", tempoConver);

            this.cmbCliente.DataSource = listaClientes;
            TimeSpan tempoCarregaCmd = stopwatch.Elapsed;
            string strCarregaCmd = tempoCarregaCmd.ToString(@"hh\:mm\:ss\.fff");
            MeuIni.WriteString("Clientes", "Carregamento em cmbCliente", strCarregaCmd);
            stopwatch.Stop();


            //Stopwatch stopwatch = new Stopwatch();
            //INI MeuIni = new INI();
            //stopwatch.Start();
            //DataTable dadosCliente = Cliente.getDadosOrdenados();
            //MeuIni.WriteString("Clientes", "Quantidade", dadosCliente.Rows.Count.ToString());

            //TimeSpan tempoDecorrido = stopwatch.Elapsed;
            //string tempoStr = tempoDecorrido.ToString(@"hh\:mm\:ss\.fff");
            //MeuIni.WriteString("Clientes", "getDadosOrdenados", tempoStr);

            //List<tb.Cliente> listaClientes = ConvertDataTableToList<tb.Cliente>(dadosCliente);
            //TimeSpan tempoConvert = stopwatch.Elapsed;
            //string tempoConver = tempoDecorrido.ToString(@"hh\:mm\:ss\.fff");
            //MeuIni.WriteString("Clientes", "ConvertDataTableToList", tempoConver);

            //this.cmbCliente.DataSource = listaClientes;
            //TimeSpan tempoCarregaCmd = stopwatch.Elapsed;
            //string strCarregaCmd = tempoDecorrido.ToString(@"hh\:mm\:ss\.fff");
            //MeuIni.WriteString("Clientes", "Carregamento em cmbCliente", strCarregaCmd);
            //stopwatch.Stop();

        }

        public List<T> ConvertDataTableToList<T>(DataTable dataTable) where T : new()
        {
            List<T> list = new List<T>();

            foreach (DataRow row in dataTable.Rows)
            {
                T item = new T();
                foreach (DataColumn column in dataTable.Columns)
                {
                    // Certifique-se de que o nome da coluna no DataTable corresponda ao nome da propriedade na classe T.
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


    }
}
