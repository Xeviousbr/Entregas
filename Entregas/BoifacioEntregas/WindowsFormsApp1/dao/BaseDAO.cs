using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        public virtual object Apagar(int direcao)
        {
            return null;
        }

        public virtual object GetUltimo()
        {
            return null;
        }

        public virtual object ParaTraz()
        {
            return null;
        }

        public virtual object ParaFrente()
        {
            return null;
        }

        public virtual object GetEsse()
        {
            return null;
        }
        

    }
}
