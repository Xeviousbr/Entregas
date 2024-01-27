﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace BonifacioEntregas.dao
{
    public class ClienteDAO : BaseDAO
    {
        // private int Linhas;

        protected int id { get; set; }
        public string Nome { get; set; }
        public string Telefone { get; set; }

        public string email { get; set; }
        public string Ender { get;  set; }

        public string NrOutro { get; set; }

        public void Addcliente(tb.Cliente cliente)
        {
            // Código para adicionar um novo cliente ao banco de dados
        }

        //public override void Grava(object obj)
        //{
        //    ClienteDAO cliente = (ClienteDAO)obj;
        //    string query;
        //    List<OleDbParameter> parameters;
        //    int result = 0;
        //    if (cliente.Adicao)
        //    {
        //        query = "INSERT INTO Clientes (NrCli, Nome, Telefone, email, Ender, NrOutro) VALUES (?, ?, ?, ?, ?, ?)";
        //        parameters = ConstruirParametroscliente(cliente, false);
        //        result = ExecutarComandoSQL(query, parameters);
        //    }
        //    else
        //    {
        //        query = $"UPDATE Clientes SET Nome = '{cliente.Nome}', Telefone = '{cliente.Telefone}', email = '{cliente.email}', Ender = '{cliente.Ender}', NrOutro = '{cliente.NrOutro}' WHERE NrCli = {cliente.id}";
        //        result = ExecutarComandoSQL(query, null);
        //    }
        //}

        public override void Grava(object obj)
        {
            ClienteDAO cliente = (ClienteDAO)obj;
            string query;
            List<OleDbParameter> parameters;
            int result = 0;
            if (cliente.Adicao)
            {
                query = "INSERT INTO Clientes (NrCli, Nome, Telefone, email, Ender, NrOutro) VALUES (?, ?, ?, ?, ?, ?)";
                parameters = ConstruirParametroscliente(cliente, true);
            }
            else
            {
                query = "UPDATE Clientes SET Nome = ?, Telefone = ?, email = ?, Ender  =?, NrOutro=? WHERE NrCli = ?";
                parameters = ConstruirParametroscliente(cliente, false);
            }

            try
            {
                gen.ExecutarComandoSQL(query, parameters);
            }
            catch (Exception ex)
            {
                string x = ex.ToString();
                MessageBox.Show(x, "Erro na operação do banco de dados", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private List<OleDbParameter> ConstruirParametroscliente(ClienteDAO cliente, bool inserindo)
        {
            int iNrOutro = 0;
            int.TryParse(cliente.NrOutro, out iNrOutro);
            var parametros = new List<OleDbParameter>
            {
                new OleDbParameter("@Nome", cliente.Nome),
                new OleDbParameter("@Telefone", cliente.Telefone),
                new OleDbParameter("@email", cliente.email),
                new OleDbParameter("@Ender", cliente.Ender),
                new OleDbParameter("@NrOutro", iNrOutro) 
            };
            if (inserindo)
            {
                parametros.Insert(0, new OleDbParameter("@NrCli", VeUltReg() + 1));
            }
            else
            {
                parametros.Add(new OleDbParameter("@NrCli", cliente.id));
            }
            return parametros;
        }

        private int VeUltReg()
        {
            string query = "SELECT Max(NrCli) as NrCli FROM Clientes";
            using (OleDbConnection connection = new OleDbConnection(gen.connectionString))
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
                                return Convert.ToInt32(reader["NrCli"]);
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

        //public int ExecutarComandoSQL(string query, List<OleDbParameter> parameters)
        //{
        //    using (OleDbConnection connection = new OleDbConnection(gen.connectionString))
        //    {
        //        using (OleDbCommand command = new OleDbCommand(query, connection))
        //        {
        //            if (parameters != null)
        //            {
        //                foreach (var param in parameters)
        //                {
        //                    command.Parameters.Add(param);
        //                }
        //            }
        //            connection.Open();
        //            return command.ExecuteNonQuery();
        //        }
        //    }
        //}

        public override tb.IDataEntity Apagar(int direcao, tb.IDataEntity entidade)
        {
            gen.ExecutarComandoSQL("DELETE FROM Clientes WHERE NrCli = " + id.ToString(), null);
            tb.Cliente proximocliente = direcao > -1 ? ParaFrente() as tb.Cliente : ParaTraz() as tb.Cliente;
            if (proximocliente == null || proximocliente.Id == 0)
            {
                proximocliente = direcao > -1 ? ParaTraz() as tb.Cliente : ParaFrente() as tb.Cliente;
            }
            return proximocliente ?? new tb.Cliente();
        }

        internal int RetIdNrAlter(string searchText)
        {
            string query = "SELECT * FROM Clientes Where NrOutro = " + searchText;
            tb.Cliente reg =ExecutarConsultacliente(query);
            if (reg==null)
            {
                return 0;
            } else
            {
                return reg.Id;
            }            
        }

        public override tb.IDataEntity GetEsse()
        {
            return (tb.Cliente)new tb.Cliente
            {
                Id = id,
                Nome = Nome,
                Telefone = Telefone,
                email = email,
                Ender=Ender,
                NrOutro= NrOutro
            };
        }

        public override object GetUltimo()
        {
            string query = "SELECT TOP 1 * FROM Clientes ORDER BY NrCli Desc";
            return ExecutarConsultacliente(query);
        }

        public override tb.IDataEntity ParaTraz()
        {
            string query = $"SELECT TOP 1 * FROM Clientes Where Nome < '{Nome}' ORDER BY Nome Desc";
            return ExecutarConsultacliente(query);
        }

        public override tb.IDataEntity ParaFrente()
        {
            string query = $"SELECT TOP 1 * FROM Clientes Where Nome > '{Nome}' ORDER BY Nome ";
            return ExecutarConsultacliente(query);
        }

        private tb.Cliente ExecutarConsultacliente(string query)
        {
            using (OleDbConnection connection = new OleDbConnection(gen.connectionString))
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
                                id = Convert.ToInt32(reader["NrCli"]);
                                Telefone = reader["Telefone"].ToString();
                                email = reader["email"].ToString();
                                Ender = reader["Ender"].ToString();
                                NrOutro = reader["NrOutro"].ToString();
                                return (tb.Cliente)GetEsse();
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

        private DataTable ExecutarConsulta(string query)
        {
            using (OleDbConnection connection = new OleDbConnection(gen.connectionString))
            {
                try
                {
                    connection.Open();
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            DataTable dataTable = new DataTable();
                            dataTable.Columns.Add("id", typeof(int));
                            dataTable.Columns.Add("Nome", typeof(string));
                            dataTable.Columns.Add("Telefone", typeof(string));
                            dataTable.Columns.Add("email", typeof(string));
                            dataTable.Columns.Add("Ender", typeof(string));
                            dataTable.Columns.Add("NrOutro", typeof(string)); 
                            while (reader.Read())
                            {
                                DataRow row = dataTable.NewRow();
                                row["id"] = reader["NrCli"];
                                row["Nome"] = reader["Nome"];
                                row["Telefone"] = reader["Telefone"];
                                row["email"] = reader["email"];
                                row["Ender"] = reader["Ender"];
                                row["NrOutro"] = reader["NrOutro"];
                                dataTable.Rows.Add(row);
                            }
                            return dataTable;
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Aqui você pode decidir como lidar com a exceção
                    throw;
                }
            }
        }

        public override DataTable getDadosOrdenados()
        {
            string query = "SELECT * FROM Clientes Order By Nome ";
            return ExecutarConsulta(query);
        }

        public override DataTable getDados()
        {
            string query = $"SELECT * FROM Clientes";
            return ExecutarConsulta(query);
        }

        public override DataTable Fitrar(string pesquisar)
        {
            string query = $"SELECT * FROM Clientes Where Nome Like '%{pesquisar}%'";
            return ExecutarConsulta(query);
        }

        public override tb.IDataEntity GetPeloID(string id)
        {
            string query = $"SELECT * FROM Clientes Where NrCli = {id} ";
            return ExecutarConsultacliente(query);
        }

    }

}
