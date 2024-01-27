using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Text;

namespace BonifacioEntregas.dao
{
    public class EntregasDAO
    {
        public EntregasDAO()
        {
            
        }

        public DataTable getDados(DateTime? DT)
        {
            StringBuilder query = new StringBuilder();
            query.Append(@"SELECT
                    e.ID as Id, 
                    e.Data, 
                    m.Nome AS MotoBoy, 
                    e.Valor, 
                    e.Desconto,
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
                    e.Obs,
                    m.codi as idBoy,
                    c.NrCli,
                    e.idForma 
                FROM 
                    ((Entregas e
                    INNER JOIN Clientes c ON c.NrCli = e.idCliente)
                    INNER JOIN Mecanicos m ON m.codi = e.idBoy)");
            if (DT.HasValue)
            {
                DateTime dataInicio = DT.Value.Date;
                DateTime dataFim = dataInicio.AddDays(1).AddTicks(-1);
                string dataInicioStr = dataInicio.ToString("MM/dd/yyyy HH:mm:ss");
                string dataFimStr = dataFim.ToString("MM/dd/yyyy HH:mm:ss");
                query.AppendFormat(" WHERE e.Data BETWEEN #{0}# AND #{1}#", dataInicioStr, dataFimStr);
            }
            query.Append(" Order By e.ID desc");
            DataTable dt = ExecutarConsulta(query.ToString());
            return dt;
        }

        private DataTable ExecutarConsulta(string query)
        {
            DataTable dataTable = new DataTable();
            using (OleDbConnection connection = new OleDbConnection(gen.connectionString))
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

        public void Adiciona(int idBoy, int idForma, float valor, int idcliente, float compra, string Obs, float desc)
        {
            String sql = @"INSERT INTO Entregas (idCliente, idForma, idBoy, Valor, VlNota, Obs, Desconto, Data) VALUES ("
                + idcliente.ToString() + ", "
                + idForma.ToString() + ", "
                + idBoy.ToString() + ", "                
                + gen.sv(valor) + ", "
                + gen.sv(compra) + ", "
                + gen.fa(Obs) + ", "
                + gen.sv(desc)
                + ",Now)";
            ExecutarComandoSQL(sql);
        }

        private void ExecutarComandoSQL(string query)
        {
            using (OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + gen.CaminhoBase + ";"))
            {
                connection.Open();
                using (OleDbCommand command = new OleDbCommand(query, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        public void Edita(int iID, int idBoy, int idForma, float valor, int idCliente, float compra, string obs, float desc)
        {
            String sql = @"UPDATE Entregas SET 
                idCliente = " + idCliente.ToString() + 
                            ",idForma = " + idForma.ToString() + 
                            ",idBoy = " + idBoy.ToString() + 
                            ",Valor = " + gen.sv(valor) + 
                            ",VlNota = " + gen.sv(compra) + 
                            ",Obs = " + gen.fa(obs) + 
                            ",Desconto = " + gen.sv(desc) +
                            " WHERE ID = " + iID.ToString();
            ExecutarComandoSQL(sql);

        }

    }
}
