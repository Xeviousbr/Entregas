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
        private string EsseCNH = "";
        private DateTime EsseDataValidadeCNH;

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
            List<OleDbParameter> parameters;
            int result = 0;
            if (entregador.Id == 0)
            {
                query = "INSERT INTO Mecanicos (codi, Oper, Nome, Telefone, CNH, DataValidadeCNH) VALUES (?, ?, ?, ?, ?, ?)";
                parameters = ConstruirParametrosEntregador(entregador, true);
            }
            else
            {
                query = "UPDATE Mecanicos SET Nome = ?, Telefone = ?, CNH = ?, DataValidadeCNH = ? WHERE codi = ?";
                parameters = ConstruirParametrosEntregador(entregador, false);
            }

            try
            {
                result = ExecutarComandoSQL(query, parameters);
            }
            catch (Exception ex)
            {
                string x = ex.ToString();
                // Considerar um melhor tratamento de exceções ou log
            }
        }

        //public void Grava(tb.Entregador entregador)
        //{
        //    string query;
        //    List<OleDbParameter> parameters;
        //    int result = 0;
        //    if (entregador.Id == 0)
        //    {
        //        Esseid = VeUltReg() + 1;
        //        query = "INSERT INTO Mecanicos (codi, Oper, Nome, Telefone, CNH, DataValidadeCNH) VALUES (?, ?, ?, ?, ?, ?)";
        //        parameters = new List<OleDbParameter>
        //        {
        //            new OleDbParameter("@codi", Esseid),
        //            new OleDbParameter("@Oper", 3),
        //            new OleDbParameter("@Nome", entregador.Nome),
        //            new OleDbParameter("@Telefone", entregador.Telefone),
        //            new OleDbParameter("@CNH", entregador.CNH),
        //            new OleDbParameter("@DataValidadeCNH", entregador.DataValidadeCNH)
        //        };
        //    }
        //    else
        //    {
        //        query = "UPDATE Mecanicos SET Nome = ?, Telefone = ?, CNH = ?, DataValidadeCNH = ? WHERE codi = ?";
        //        parameters = new List<OleDbParameter>
        //        {
        //            new OleDbParameter("@Nome", entregador.Nome),
        //            new OleDbParameter("@Telefone", entregador.Telefone),
        //            new OleDbParameter("@CNH", entregador.CNH),
        //            new OleDbParameter("@DataValidadeCNH", entregador.DataValidadeCNH),
        //            new OleDbParameter("@codi", entregador.Id)
        //        };
        //    }
        //    try
        //    {
        //        result = ExecutarComandoSQL(query, parameters);
        //    }
        //    catch (Exception ex)
        //    {
        //        string x = ex.ToString();
        //    }
        //}

        private List<OleDbParameter> ConstruirParametrosEntregador(tb.Entregador entregador, bool inserindo)
        {
            var parametros = new List<OleDbParameter>
            {
                new OleDbParameter("@Nome", entregador.Nome),
                new OleDbParameter("@Telefone", entregador.Telefone),
                new OleDbParameter("@CNH", entregador.CNH),
                new OleDbParameter("@DataValidadeCNH", entregador.DataValidadeCNH)
            };
            if (inserindo)
            {
                parametros.Insert(0, new OleDbParameter("@Oper", 3));
                parametros.Insert(0, new OleDbParameter("@codi", VeUltReg() + 1));
            }
            else
            {
                parametros.Add(new OleDbParameter("@codi", entregador.Id));
            }
            return parametros;
        }


        //public void Grava(tb.Entregador entregador)
        //{
        //    string query;
        //    int result = 0;
        //    List<OleDbParameter> parameters;
        //    if (entregador.Id == 0)
        //    {
        //        Esseid = VeUltReg()+1;
        //        query = "INSERT INTO Mecanicos (codi, Oper, Nome, Telefone) VALUES (?, ?, ?, ?)";
        //        parameters = new List<OleDbParameter>
        //        {
        //            new OleDbParameter("@codi", Esseid),
        //            new OleDbParameter("@Oper", 3),
        //            new OleDbParameter("@Nome", entregador.Nome),
        //            new OleDbParameter("@Telefone", entregador.Telefone)
        //        };
        //    }
        //    else
        //    {
        //        query = "UPDATE Mecanicos SET Nome = ?, Telefone = ? WHERE codi = ?";
        //        parameters = new List<OleDbParameter>
        //        {
        //            new OleDbParameter("@Nome", entregador.Nome),
        //            new OleDbParameter("@Telefone", entregador.Telefone),
        //            new OleDbParameter("@codi", entregador.Id)
        //        };
        //    }
        //    try
        //    {
        //        result = ExecutarComandoSQL(query, parameters);
        //    }
        //    catch (Exception ex)
        //    {
        //        string x = ex.ToString();
        //    }
        //}

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

        public tb.Entregador Apagar(int direcao)
        {
            ExecutarComandoSQL("DELETE FROM Mecanicos WHERE codi = " + Esseid.ToString(), null);
            tb.Entregador proximoEntregador = direcao > -1 ? ParaFrente() : ParaTraz();
            if (proximoEntregador == null || proximoEntregador.Id == 0)
            {
                proximoEntregador = direcao > -1 ? ParaTraz() : ParaFrente();
            }
            return proximoEntregador ?? new tb.Entregador();
        }

        public tb.Entregador GetEsse()
        {
            return new tb.Entregador
            {
                Id = Esseid,
                Nome = EsseNome,
                Telefone = EsseTelefone,
                CNH = EsseCNH,
                DataValidadeCNH= EsseDataValidadeCNH
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
                                EsseCNH = reader["CNH"].ToString();
                                if (reader["DataValidadeCNH"] != DBNull.Value)
                                {
                                    EsseDataValidadeCNH = Convert.ToDateTime(reader["DataValidadeCNH"]);
                                }
                                else
                                {
                                    EsseDataValidadeCNH = DateTime.MinValue; 
                                }
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
