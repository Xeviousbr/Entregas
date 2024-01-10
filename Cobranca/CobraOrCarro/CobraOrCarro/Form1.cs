using System;
using System.Data.OleDb;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CobraOrCarro
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Prog\OrCarro\Cobranca\OrCarro.mdb";
            string ret = "";
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();

                    // Exemplo de comando SQL para inserir dados
                    string commandString = "SELECT UtComissoes FROM Config";

                    using (OleDbCommand command = new OleDbCommand(commandString, connection))
                    {
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                // Supondo que 'UtComissoes' é um tipo de dado numérico ou de texto
                                ret = reader["UtComissoes"].ToString();
                                Console.WriteLine(ret);
                            }
                        }
                    }
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                // Tratamento de exceções
                Console.WriteLine(ex.Message);
            }

        }
    }
}

//SELECT
//    Clientes.Nome AS Nome,
//    Clientes.Telefone AS Telefone,
//    DATEDIFF(DAY, CONVERT(DATE, GETDATE()), CONVERT(DATE, Orcamento.Data)) AS Dias
//FROM
//    Orcamento
//INNER JOIN
//    Clientes
//    ON Orcamento.Cliente = Clientes.Nome
//WHERE
//    Orcamento.Pagamento = '30/12/1899 00:00:00'