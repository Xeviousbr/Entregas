using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BonifacioEntregas.dao
{
    public class EntregadorDAO
    {
        private string connectionString;

        public EntregadorDAO()
        {
            this.connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalConfig.CaminhoBase + ";";
        }

        public List<tb.Entregador> GetAllEntregadores()
        {
            List<tb.Entregador> entregadores = new List<tb.Entregador>();

            // Código para buscar todos os entregadores do banco de dados
            // Utilizando a conexão com o banco e comando SQL

            return entregadores;
        }

        public void AddEntregador(tb.Entregador entregador)
        {
            // Código para adicionar um novo entregador ao banco de dados
        }

        public tb.Entregador GetUltimoEntregador()
        {
            using (OleDbConnection connection = new OleDbConnection(this.connectionString))
            {
                try
                {
                    connection.Open();
                    string query = "SELECT TOP 1 * FROM Mecanicos ORDER BY codi Desc "; // Substitua com o nome correto da tabela e coluna
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                return new tb.Entregador
                                {
                                    Id = Convert.ToInt32(reader["codi"]),
                                    Nome = reader["Nome"].ToString(),
                                    Telefone = reader["Telefone"].ToString(),
                                    // Outras propriedades...
                                };
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Tratamento de exceções adequado
                    string x = ex.ToString();
                    throw;
                }
            }
            return null; // Ou lance uma exceção se preferir
        }

    }

}
