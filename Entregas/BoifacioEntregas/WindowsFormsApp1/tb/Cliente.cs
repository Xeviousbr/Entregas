namespace BonifacioEntregas.tb
{
    public class Cliente : IDataEntity
    {
        public int Id { get; set; }
        public bool Adicao { get; set; }

        //[CampoTag("O")]
        public string Nome { get; set; }

        //[CampoTag("O")]
        public string Telefone { get; set; }

        public string email { get; set; }

        //[CampoTag("O")]
        public string Ender { get; set; }
    }
}
