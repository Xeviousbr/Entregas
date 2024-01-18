﻿using System;
using System.Collections.Generic;
using System.Data.OleDb;

namespace BonifacioEntregas.dao
{
    public class EntregadorDAO: BaseDAO  // public EntregadorDAO() : base()
    {
        protected int id { get; set; }
        public string Nome { get; set; }
        public string Telefone { get; set; }
        public string CNH { get; set; }
        public DateTime DataValidadeCNH { get; set; }

        public EntregadorDAO()
        {
            
        }

        public void AddEntregador(tb.Entregador entregador)
        {
            // Código para adicionar um novo entregador ao banco de dados
        }

        public override void Grava(object obj)
        {
            EntregadorDAO entregador = (EntregadorDAO)obj;
            string query;
            List<OleDbParameter> parameters;
            int result = 0;
            if (entregador.Adicao)
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

        private List<OleDbParameter> ConstruirParametrosEntregador(EntregadorDAO entregador, bool inserindo)
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
                parametros.Add(new OleDbParameter("@codi", entregador.id)); 
            }
            return parametros;
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

        public override tb.IDataEntity Apagar(int direcao, tb.IDataEntity entidade)
        {
            ExecutarComandoSQL("DELETE FROM Mecanicos WHERE codi = " + id.ToString(), null);
            tb.Entregador proximocliente = direcao > -1 ? ParaFrente() as tb.Entregador : ParaTraz() as tb.Entregador;
            if (proximocliente == null || proximocliente.Id == 0)
            {
                proximocliente = direcao > -1 ? ParaTraz() as tb.Entregador : ParaFrente() as tb.Entregador;
            }
            return proximocliente ?? new tb.Entregador();
        }

        public override tb.IDataEntity GetEsse()
        {
            return (tb.Entregador)new tb.Entregador
            {
                Id = id,
                Nome = Nome,
                Telefone = Telefone,
                CNH = CNH,
                DataValidadeCNH = DataValidadeCNH
            };

        }

        public override object GetUltimo()
        {
            string query = "SELECT TOP 1 * FROM Mecanicos Where Oper = 3 ORDER BY codi Desc"; 
            return ExecutarConsultaEntregador(query);
        }

        public override tb.IDataEntity ParaTraz()
        {
            string query = $"SELECT TOP 1 * FROM Mecanicos Where Oper = 3 and Nome < '{Nome}' ORDER BY Nome Desc"; 
            return ExecutarConsultaEntregador(query);
        }

        public override tb.IDataEntity ParaFrente()
        {
            string query = $"SELECT TOP 1 * FROM Mecanicos Where Oper = 3 and Nome > '{Nome}' ORDER BY Nome ";
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
                                Nome = reader["Nome"].ToString();
                                id = Convert.ToInt32(reader["codi"]);
                                Telefone = reader["Telefone"].ToString();
                                CNH = reader["CNH"].ToString();
                                if (reader["DataValidadeCNH"] != DBNull.Value)
                                {
                                    DataValidadeCNH = Convert.ToDateTime(reader["DataValidadeCNH"]);
                                }
                                else
                                {
                                    DataValidadeCNH = DateTime.MinValue; 
                                }
                                return (tb.Entregador)GetEsse();
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
