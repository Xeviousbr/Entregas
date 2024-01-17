using System;
using System.Collections.Generic;
using System.Data.OleDb;

namespace BonifacioEntregas.dao
{
    public class ClienteDAO : BaseDAO
    {

        protected int id { get; set; }
        public string Nome { get; set; }
        public string Telefone { get; set; }

        public string email { get; set; }
        public string Ender { get;  set; }

        public void Addcliente(tb.Cliente cliente)
        {
            // Código para adicionar um novo cliente ao banco de dados
        }

        public override void Grava(object obj)
        {
            ClienteDAO cliente = (ClienteDAO)obj;
            string query;
            List<OleDbParameter> parameters;
            int result = 0;
            if (cliente.Adicao)
            {
                query = "INSERT INTO Clientes (NrCli, Nome, Telefone, email, Ender) VALUES (?, ?, ?, ?, ?)";
                parameters = ConstruirParametroscliente(cliente, true);
            }
            else
            {
                query = "UPDATE Clientes SET Nome = ?, Telefone = ?, email = ?, Ender  =? WHERE NrCli = ?";
                parameters = ConstruirParametroscliente(cliente, false);
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

        private List<OleDbParameter> ConstruirParametroscliente(ClienteDAO cliente, bool inserindo)
        {
            var parametros = new List<OleDbParameter>
            {
                new OleDbParameter("@Nome", cliente.Nome),
                new OleDbParameter("@Telefone", cliente.Telefone),
                new OleDbParameter("@email", cliente.email),
                new OleDbParameter("@Ender", cliente.Ender),
                
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
            string query = $"SELECT Max(NrCli) as NrCli FROM Clientes";
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

        public override object Apagar(int direcao)
        {
            ExecutarComandoSQL("DELETE FROM Clientes WHERE NrCli = " + id.ToString(), null);
            tb.Cliente proximocliente = direcao > -1 ? ParaFrente() as tb.Cliente : ParaTraz() as tb.Cliente;
            if (proximocliente == null || proximocliente.Id == 0)
            {
                proximocliente = direcao > -1 ? ParaTraz() as tb.Cliente : ParaFrente() as tb.Cliente;
            }
            return proximocliente ?? new tb.Cliente();
        }

        public override object GetEsse()
        {
            return (tb.Cliente)new tb.Cliente
            {
                Id = id,
                Nome = Nome,
                Telefone = Telefone,
                email = email,
                Ender=Ender
            };

        }

        public override object GetUltimo()
        {
            // string query = "SELECT TOP 1 * FROM Clientes ORDER BY NrCli Desc";
            string query = "SELECT TOP 1 * FROM Clientes Where NrCli =4925 ";
            return ExecutarConsultacliente(query);
        }

        public override object ParaTraz()
        {
            string query = $"SELECT TOP 1 * FROM Clientes Where Nome < '{Nome}' ORDER BY Nome Desc";
            return ExecutarConsultacliente(query);
        }

        public override object ParaFrente()
        {
            string query = $"SELECT TOP 1 * FROM Clientes Where Nome > '{Nome}' ORDER BY Nome ";
            return ExecutarConsultacliente(query);
        }

        private tb.Cliente ExecutarConsultacliente(string query)
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
                                id = Convert.ToInt32(reader["NrCli"]);
                                Telefone = reader["Telefone"].ToString();
                                email = reader["email"].ToString();
                                Ender = reader["Ender"].ToString();
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

    }

}
