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
        private string EsseTelefone = "";

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

        public void Grava(tb.Entregador entregador)
        {
            string query = "UPDATE Mecanicos SET Nome = ?, Telefone = ? WHERE codi = ?";
            List<OleDbParameter> parameters = new List<OleDbParameter>
            {
                new OleDbParameter("@Nome", entregador.Nome),
                new OleDbParameter("@Telefone", entregador.Telefone),
                new OleDbParameter("@codi", Esseid) // Garanta que Esseid esteja definido corretamente
            };

            try
            {
                int result = ExecutarComandoSQL(query, parameters);
                // Tratar o resultado conforme necessário
            }
            catch (Exception ex)
            {
                // Tratamento de erro
                // throw ou outra lógica de erro
            }
        }

        public int ExecutarComandoSQL(string query, List<OleDbParameter> parameters)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                using (OleDbCommand command = new OleDbCommand(query, connection))
                {
                    foreach (var param in parameters)
                    {
                        command.Parameters.Add(param);
                    }

                    connection.Open();
                    return command.ExecuteNonQuery();
                }
            }
        }

        public void Apagar()
        {
            ExecutarComandoSQL($"Delete From Mecanicos Where codi = {Esseid} ", null);
        }

        public tb.Entregador GetEsse()
        {
            return new tb.Entregador
            {
                Id = Esseid,
                Nome = EsseNome,
                Telefone = EsseTelefone,
            };
        }

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
                                EsseTelefone = reader["Telefone"].ToString();
                                return GetEsse();
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
