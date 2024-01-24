using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;

namespace BonifacioEntregas.dao
{
    public class EntregasDAO
    {
        private string connectionString;

        public EntregasDAO()
        {
            this.connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalConfig.CaminhoBase + ";";
        }

        public DataTable getDados()
        {
            string query = @"SELECT
                                e.ID as Id, 
                                e.Data, 
                                m.Nome AS MotoBoy, 
                                e.Valor, 
                                SWITCH(
                                    e.idForma = 0, 'Anotado',
                                    e.idForma = 1, 'Cartão',
                                    e.idForma = 2, 'Dinheiro',
                                    e.idForma = 3, 'Pix',
                                    e.idForma = 5, 'Troca',
                                    TRUE, 'Desconhecido'
                                ) AS Pagamento,
                                e.VlNota as Compra, 
                                c.Nome AS Cliente,
                                e.Obs 
                            FROM 
                                ((Entregas e
                                INNER JOIN Clientes c ON c.NrCli = e.idCliente)
                                INNER JOIN Mecanicos m ON m.codi = e.idBoy)
                            Order By e.ID desc ";
            DataTable dt = ExecutarConsulta(query);
            return dt;
        }

        private DataTable ExecutarConsulta(string query)
        {
            DataTable dataTable = new DataTable();
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection))
                    {
                        adapter.Fill(dataTable);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            return dataTable;
        }

        public void Adiciona(int idBoy, int idForma, float valor, int idCliente, float compra, string Obs)
        {
            String sql = @"INSERT INTO Entregas (idCliente, idForma, idBoy, Valor, VlNota, Obs, Data) VALUES ("                 
                + idCliente.ToString() + ", " 
                + idForma.ToString() + ", " 
                + idBoy.ToString()+ " ,"
                + valor.ToString()+ ", "
                + compra.ToString() + ", "
                + "'" + Obs +"'" +
                ",Now) ";
            ExecutarComandoSQL(sql);
        }

        private void ExecutarComandoSQL(string query)
        {
            using (OleDbConnection connection = new OleDbConnection(this.connectionString))
            {
                connection.Open();
                using (OleDbCommand command = new OleDbCommand(query, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

    }
}
