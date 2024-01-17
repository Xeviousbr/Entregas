using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BonifacioEntregas.tb
{
    public class Cliente : IDataEntity
    {
        public int Id { get; set; }
        public bool Adicao { get; set; }

        public string Nome { get; set; }
        public string Telefone { get; set; }

        public string email { get; set; }
    }
}
