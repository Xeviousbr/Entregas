

using System;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;

namespace AuxSql
{
    public partial class Form1 : Form
    {

        private string connectionString = "";

        public Form1()
        {
            InitializeComponent();
            string databasePath = EncontrarPrimeiroMDB(AppDomain.CurrentDomain.BaseDirectory);

            if (!string.IsNullOrEmpty(databasePath))
            {
                connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + databasePath + ";";
            }
            else
            {
                MessageBox.Show("Nenhum arquivo .mdb encontrado na pasta do executável.");
            }
        }

        private string EncontrarPrimeiroMDB(string directoryPath)
        {
            string[] mdbFiles = Directory.GetFiles(directoryPath, "*.mdb", SearchOption.TopDirectoryOnly);

            if (mdbFiles.Length > 0)
            {
                return mdbFiles[0]; 
            }
            else
            {
                return null;
            }
        }

        private void ExecuteQuery(string query)
        {
            try
            {
                using (OleDbConnection connection = new OleDbConnection(this.connectionString))
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        int affectedRows = command.ExecuteNonQuery();
                        MessageBox.Show($"{affectedRows} linhas afetadas.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao executar a consulta: " + ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExecuteQuery(textBoxQuery.Text);
        }
    }

}
