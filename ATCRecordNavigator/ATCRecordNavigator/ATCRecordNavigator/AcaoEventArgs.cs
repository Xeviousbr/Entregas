using System;

public class AcaoEventArgs : EventArgs
{
    public string Acao { get; set; }
    public string Dado { get; set; } // Use um tipo apropriado para o ID

    public AcaoEventArgs(string acao, string bscDado = "")
    {
        Acao = acao;
        Dado = bscDado;
    }
}
