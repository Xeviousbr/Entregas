using System;
using System.Collections.Generic;
using System.Data.OleDb;

namespace BonifacioEntregas.dao
{
    public class EntregadorDAO
    {
        private string connectionString;
        public int Esseid = 0;
        private string EsseNome = "";
        private string EsseTelefone = "";

        public EntregadorDAO()
        {
            this.connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalConfig.CaminhoBase + ";";
        }

        public void AddEntregador(tb.Entregador entregador)
        {
            // Código para adicionar um novo entregador ao banco de dados
        }

        public void Grava(tb.Entregador entregador)
        {
            string query;
            int result = 0;
            List<OleDbParameter> parameters;
            if (entregador.Id == 0)
            {
                Esseid = VeUltReg()+1;
                query = "INSERT INTO Mecanicos (codi, Oper, Nome, Telefone) VALUES (?, ?, ?, ?)";
                parameters = new List<OleDbParameter>
                {
                    new OleDbParameter("@codi", Esseid),
                    new OleDbParameter("@Oper", 3),
                    new OleDbParameter("@Nome", entregador.Nome),
                    new OleDbParameter("@Telefone", entregador.Telefone)
                };
            }
            else
            {
                query = "UPDATE Mecanicos SET Nome = ?, Telefone = ? WHERE codi = ?";
                parameters = new List<OleDbParameter>
                {
                    new OleDbParameter("@Nome", entregador.Nome),
                    new OleDbParameter("@Telefone", entregador.Telefone),
                    new OleDbParameter("@codi", entregador.Id)
                };
            }
            try
            {
                result = ExecutarComandoSQL(query, parameters);
            }
            catch (Exception ex)
            {
                string x = ex.ToString();
            }
        }

        private int VeUltReg()
        {
            string query = $"SELECT Max(codi) as codi FROM Mecanicos";
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
                                return Convert.ToInt32(reader["codi"]);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    string x = ex.ToString();
                }
                return 0;
            }
        }

        public int ExecutarComandoSQL(string query, List<OleDbParameter> parameters)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                using (OleDbCommand command = new OleDbCommand(query, connection))
                {
                    if (parameters != null)
                    {
                        foreach (var param in parameters)
                        {
                            command.Parameters.Add(param);
                        }
                    }
                    connection.Open();
                    return command.ExecuteNonQuery();
                }
            }
        }

        public tb.Entregador Apagar(int Direcao)
        {
            ExecutarComandoSQL("Delete From Mecanicos Where codi = "+Esseid.ToString(),null);
            string query = "";
            if (Direcao>-1)
            {
                query = "SELECT TOP 1 * FROM Mecanicos Where codi<" + Esseid.ToString()+" ORDER BY codi Desc";
            } else
            {
                query = "SELECT TOP 1 * FROM Mecanicos Where codi>" + Esseid.ToString() + " ORDER BY codi";
            }
            return ExecutarConsultaEntregador(query);            
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
            string query = "SELECT TOP 1 * FROM Mecanicos Where Oper = 3 ORDER BY codi Desc"; 
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
