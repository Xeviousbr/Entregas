using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace BonifacioEntregas.dao
{
    public class VendedoresDAO : BaseDAO
    {
        public int Id { get; set; }

        public string Nome { get; set; }

        public string Loja { get; set; }

        public VendedoresDAO()
        {
            
        }

        //public DataTable GetAllVendedores()
        //{
        //    string query = "SELECT * FROM Vendedores";
        //    return ExecutarConsulta(query);
        //}

        //private DataTable ExecutarConsulta(string query)
        //{
        //    DataTable dataTable = new DataTable();
        //    using (OleDbConnection connection = new OleDbConnection(gen.connectionString))
        //    {
        //        try
        //        {
        //            connection.Open();
        //            using (OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection))
        //            {
        //                adapter.Fill(dataTable);
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            Console.WriteLine(ex.Message);
        //        }
        //    }
        //    return dataTable;
        //}

        public void AdicionaVendedor(string nome, string loja)
        {
            String sql = @"INSERT INTO Vendedores (Nome, Loja) VALUES ('"
                + gen.fa(nome) + "', '"
                + gen.fa(loja) + "')";
            gen.ExecutarComandoSQL(sql);
        }

        public void EditaVendedor(int id, string nome, string loja)
        {
            String sql = @"UPDATE Vendedores SET 
                Nome = '" + gen.fa(nome) +
                "', Loja = '" + gen.fa(loja) +
                "' WHERE ID = " + id.ToString();
            gen.ExecutarComandoSQL(sql);
        }
        public tb.Vendedor GetUltimoVendedor()
        {
            string query = "SELECT TOP 1 * FROM Vendedores ORDER BY ID Desc";
            tb.Vendedor X = ExecutarConsultaVendedor(query);
            return X;
        }

        public override void Grava(object obj)
        {
            VendedoresDAO vendedor = (VendedoresDAO)obj;
            string query;
            List<OleDbParameter> parameters;
            int result = 0;
            if (vendedor.Adicao)
            {
                query = "INSERT INTO Vendedores (Nome, Loja) VALUES (?, ?)";
                parameters = ConstruirParametro(vendedor, true);
            }
            else
            {
                query = "UPDATE Vendedores SET Nome = ?, Loja = ? WHERE ID = ?";
                parameters = ConstruirParametro(vendedor, false);
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

        private List<OleDbParameter> ConstruirParametro(VendedoresDAO vendedor, bool inserindo)
        {

            var parametros = new List<OleDbParameter>
            {
                new OleDbParameter("@Nome", vendedor.Nome),
                new OleDbParameter("@Loja", vendedor.Loja)
            };
            if (!inserindo)
            {
                parametros.Add(new OleDbParameter("@ID", vendedor.Id));
            }
            return parametros;
        }

        public tb.Vendedor ExecutarConsultaVendedor(string query)
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
                                Id = Convert.ToInt32(reader["ID"]);
                                Nome = reader["Nome"].ToString();
                                Loja = reader["Loja"].ToString();
                                return (tb.Vendedor)GetEsse();
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

        //public override tb.IDataEntity GetEsse()
        //{
        //    return (tb.Vendedor)new tb.Vendedor
        //    {
        //        Id = Id,
        //        Nome = Nome,
        //        Loja = Loja
        //    };
        //}

    }
}
