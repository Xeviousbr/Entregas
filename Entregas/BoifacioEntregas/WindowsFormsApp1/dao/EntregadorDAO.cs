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
        private int Esseid = 0;
        private string EsseNome = "";

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

        //public tb.Entregador GetUltimoEntregador()
        //{
        //    using (OleDbConnection connection = new OleDbConnection(this.connectionString))
        //    {
        //        try
        //        {
        //            connection.Open();
        //            string query = "SELECT TOP 1 * FROM Mecanicos ORDER BY codi Desc "; // Substitua com o nome correto da tabela e coluna
        //            using (OleDbCommand command = new OleDbCommand(query, connection))
        //            {
        //                using (OleDbDataReader reader = command.ExecuteReader())
        //                {
        //                    if (reader.Read())
        //                    {
        //                        Esseid = Convert.ToInt32(reader["codi"]);
        //                        return new tb.Entregador
        //                        {
        //                            Id = Esseid,
        //                            Nome = reader["Nome"].ToString(),
        //                            Telefone = reader["Telefone"].ToString(),                                    
        //                        };
        //                    }
        //                }
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            // Tratamento de exceções adequado
        //            string x = ex.ToString();
        //            throw;
        //        }
        //    }
        //    return null; 
        //}

        public tb.Entregador GetUltimoEntregador()
        {
            string query = "SELECT TOP 1 * FROM Mecanicos Where Oper = 3 ORDER BY codi Desc"; // Ajuste a query conforme necessário
            return ExecutarConsultaEntregador(query);
        }

        public tb.Entregador ParaTraz()
        {
            string query = $"SELECT TOP 1 * FROM Mecanicos Where Oper = 3 and Nome < '{EsseNome}' ORDER BY Nome Desc"; 
            return ExecutarConsultaEntregador(query);
        }

        public tb.Entregador ParaFrente()
        {
            string query = $"SELECT TOP 1 * FROM Mecanicos Where Oper = 3 and Nome > '{EsseNome}' ORDER BY Nome ";
            return ExecutarConsultaEntregador(query);
        }

        private tb.Entregador ExecutarConsultaEntregador(string query)
        {
            using (OleDbConnection connection = new OleDbConnection(this.connectionString))
            {
                try
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                EsseNome = reader["Nome"].ToString();
                                Esseid = Convert.ToInt32(reader["codi"]);
                                string Oper = reader["Oper"].ToString();
                                return new tb.Entregador
                                {
                                    Id = Esseid,
                                    Nome = EsseNome,
                                    Telefone = reader["Telefone"].ToString(),
                                };
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Tratamento de exceções adequado
                    throw;
                }
            }
            return null;
        }

    }

}
