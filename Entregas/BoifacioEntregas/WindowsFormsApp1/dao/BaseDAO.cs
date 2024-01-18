namespace BonifacioEntregas.dao
{
    public class BaseDAO
    {
        protected string connectionString;

        public bool Adicao { get; set; }

        protected BaseDAO()
        {
            this.connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GlobalConfig.CaminhoBase + ";";
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
        

    }
}
