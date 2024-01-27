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
            string valorLimpo = new string(valorTexto.Where(c => char.IsDigit(c) || c == ',' || c == '.').ToArray());
            char decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator[0];
            if (valorLimpo.Contains('.') && valorLimpo.Contains(','))
            {
                valorLimpo = valorLimpo.Replace(".", decimalSeparator.ToString());
            }
            else if (valorLimpo.Contains('.') || valorLimpo.Contains(','))
            {
                valorLimpo = valorLimpo.Replace(',', decimalSeparator).Replace('.', decimalSeparator);
            }
            if (float.TryParse(valorLimpo, out float valorFloat))
            {
                return valorFloat;
            }
            else
            {
                return 0.0f;
            }
        }

        public static string fmtVlr(string input)
        {
            string cleanValue = new string(input.Where(c => char.IsDigit(c) || c == ',' || c == '.').ToArray());

            // Verifique se o valor tem um ponto decimal ou vírgula decimal
            char decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator[0];
            if (cleanValue.Contains('.') && cleanValue.Contains(','))
            {
                // Se houver tanto ponto quanto vírgula, use o separador decimal atual
                cleanValue = cleanValue.Replace(".", decimalSeparator.ToString());
            }
            else if (cleanValue.Contains('.') || cleanValue.Contains(','))
            {
                // Se houver apenas ponto ou apenas vírgula, substitua pelo separador decimal atual
                cleanValue = cleanValue.Replace(',', decimalSeparator).Replace('.', decimalSeparator);
            }

            // Converta o valor limpo para um valor decimal
            if (decimal.TryParse(cleanValue, out decimal value))
            {
                // Se o valor for zero, retorne uma string vazia
                if (value == 0)
                {
                    return "";
                }

                // Formate o valor decimal como uma string sem cifrão
                return value.ToString("0.00"); // "0.00" é usado para garantir duas casas decimais
            }
            else
            {
                // Se a conversão falhar, retorne a string original
                return input;
            }
        }

        public static int ConvOjbInt(object obj)
        {
            string str = obj.ToString();
            int i = int.Parse(str);
            return i;
        }

        public static string ConvOjbStr(object obj)
        {
            string str = obj.ToString();
            return str;
        }

        public static string fa(string str)
        {
            return "'" + str + "'";
        }

    }
}