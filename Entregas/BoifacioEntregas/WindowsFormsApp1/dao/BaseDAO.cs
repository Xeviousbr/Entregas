using System;
using System.Data;

namespace BonifacioEntregas.dao
{
    public abstract class BaseDAO
    {
        protected string connectionString;

        public bool Adicao { get; set; }

        protected BaseDAO()
        {
            this.connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + gen.CaminhoBase + ";";
        }

        public virtual void Grava(object obj)
        {

        }
        public virtual tb.IDataEntity Apagar(int direcao, tb.IDataEntity entidade)
        {
            return null;
        }

        public virtual object GetUltimo()
        {
            return null;
        }

        public virtual tb.IDataEntity ParaTraz()
        {
            return null;
        }

        public virtual tb.IDataEntity ParaFrente()
        {
            return null;
        }

        public virtual tb.IDataEntity GetEsse()
        {
            return null;
        }

        public virtual System.Data.DataTable CarregarDados()
        {
            return null;
        }

        public virtual DataTable getDados()
        {
            return null;
        }

        public virtual void SetarLinhas(int v)
        {

        }
        public virtual tb.IDataEntity GetPeloID(string id)
        {
            return null;
        }

        public virtual DataTable Fitrar(string pesquisar)
        {
            return null;
        }

        public virtual DataTable getDadosOrdenados()
        {
            return null;
        }

    }
}
