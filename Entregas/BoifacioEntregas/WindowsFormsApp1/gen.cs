using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BonifacioEntregas
{
    public static class gen
    {
        public static string CaminhoBase { get; set; }

        public static float LeValor(string valorTexto)
        {
            // Remova todos os caracteres que não são dígitos, pontos ou vírgulas
            string valorLimpo = new string(valorTexto.Where(c => char.IsDigit(c) || c == ',' || c == '.').ToArray());

            // Verifique se o valor tem um ponto decimal ou vírgula decimal
            char decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator[0];
            if (valorLimpo.Contains('.') && valorLimpo.Contains(','))
            {
                // Se houver tanto ponto quanto vírgula, use o separador decimal atual
                valorLimpo = valorLimpo.Replace(".", decimalSeparator.ToString());
            }
            else if (valorLimpo.Contains('.') || valorLimpo.Contains(','))
            {
                // Se houver apenas ponto ou apenas vírgula, substitua pelo separador decimal atual
                valorLimpo = valorLimpo.Replace(',', decimalSeparator).Replace('.', decimalSeparator);
            }

            // Converta o valor limpo para um valor float
            if (float.TryParse(valorLimpo, out float valorFloat))
            {
                return valorFloat;
            }
            else
            {
                // Se a conversão falhar, retorne 0 ou outro valor padrão
                return 0.0f;
            }
        }
    }
}