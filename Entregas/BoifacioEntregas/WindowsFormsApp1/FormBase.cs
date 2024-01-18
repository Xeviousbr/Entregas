using BonifacioEntregas.tb;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BonifacioEntregas
{
    public partial class FormBase : Form
    {

        protected int Direcao = 0;
        protected bool EmAdicao = false;
        protected bool Mostrando = false;
        protected dao.BaseDAO DAO;
        protected tb.IDataEntity reg;

        public FormBase()
        {
            InitializeComponent();
        }

        #region Campos

        protected bool Mostra()
        {
            if (reg == null)
            {
                return false;
            }
            else
            {
                Mostrando = true;
                foreach (Control control in this.Controls)
                {
                    if (control is TextBox textBox)
                    {
                        ProcessarTextBox(textBox);
                    }
                    else if (control is DateTimePicker dateTimePicker)
                    {
                        ProcessarDateTimePicker(dateTimePicker);
                    }
                }
                Mostrando = false;
                return true;
            }
        }

        protected void ProcessarTextBox(TextBox textBox)
        {
            string propertyName = textBox.Name.Substring(3); // Remove o prefixo 'txt'
            PropertyInfo propertyInfo = reg.GetType().GetProperty(propertyName);
            if (propertyInfo != null)
            {
                string valor = propertyInfo.GetValue(reg, null)?.ToString() ?? string.Empty;
                textBox.Text = valor;
            }
        }

        private void ProcessarDateTimePicker(DateTimePicker dtpControl)
        {
            string propertyName = dtpControl.Name.Substring(3); 
            PropertyInfo propertyInfo = reg.GetType().GetProperty(propertyName);
            if (propertyInfo != null && propertyInfo.PropertyType == typeof(DateTime) || propertyInfo.PropertyType == typeof(DateTime?))
            {
                DateTime? data = propertyInfo.GetValue(reg, null) as DateTime?;
                if (!data.HasValue || data.Value == DateTime.MinValue)
                {
                    dtpControl.CustomFormat = " ";
                    dtpControl.Format = DateTimePickerFormat.Custom;
                }
                else
                {
                    dtpControl.Value = data.Value;
                    dtpControl.Format = DateTimePickerFormat.Short;
                }
            }
        }

        private void MapearCamposParaModelo(dao.BaseDAO reg)
        {
            foreach (Control control in this.Controls)
            {
                try
                {
                    if (control is TextBox textBox)
                    {
                        MapearTextBoxParaModelo(textBox, reg);
                    }
                    else if (control is DateTimePicker dateTimePicker)
                    {
                        MapearDateTimePickerParaModelo(dateTimePicker, reg);
                    }
                }
                catch (Exception ex)
                {
                    string x = ex.ToString();
                    // Tratamento adequado de exceções
                }
            }
        }

        private void MapearTextBoxParaModelo(TextBox textBox, dao.BaseDAO reg)
        {
            string propertyName = textBox.Name.Substring(3); // Remove o prefixo 'txt'
            PropertyInfo propertyInfo = reg.GetType().GetProperty(propertyName);
            if (propertyInfo == null)
            {
                propertyInfo.SetValue(reg, null, null);
            } else
            {
                propertyInfo.SetValue(reg, textBox.Text, null);
            }
        }

        private void MapearDateTimePickerParaModelo(DateTimePicker dtpControl, dao.BaseDAO reg)
        {
            string propertyName = dtpControl.Name.Substring(3); // Remove o prefixo 'dtp'
            PropertyInfo propertyInfo = reg.GetType().GetProperty(propertyName);
            if (propertyInfo != null && (propertyInfo.PropertyType == typeof(DateTime) || propertyInfo.PropertyType == typeof(DateTime?)))
            {
                if (dtpControl.Format != DateTimePickerFormat.Custom)
                {
                    propertyInfo.SetValue(reg, dtpControl.Value, null);
                }
                else
                {
                    propertyInfo.SetValue(reg, null, null); // ou DateTime.MinValue
                }
            }
        }

        #endregion

        #region TratamentoDeTela
        protected void LimparCampos()
        {
            foreach (Control control in this.Controls)
            {
                if (control is TextBox)
                {
                    control.Text = string.Empty;
                }
            }
        }

        public void ResetarAparenciaControles()
        {
            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is TextBox)
                {
                    ctrl.BackColor = SystemColors.Window; // Cor normal
                }
                // Adicione mais lógica aqui se houver outros tipos de controles
            }
        }

        #endregion

        #region Crud
        protected void Grava(List<CampoTagInfo> tagsDosCampos)
        {
            MapearCamposParaModelo(DAO);
            List<string> criticas = FazerCriticas(DAO, tagsDosCampos);            
            if (criticas.Count == 0)
            {
                DAO.Adicao = EmAdicao;                
                DAO.Grava(DAO);
                EmAdicao = false;
                cntrole1.ControlesNormais();
            }
            else
            {                
                string mensagemCritica = string.Join("\n", criticas);
                MessageBox.Show(mensagemCritica, "Críticas", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        #endregion

        #region Criticas

        public List<string> FazerCriticas<T>(T objeto, List<CampoTagInfo> tagsDosCampos) where T : class
        {
            List<string> criticas = new List<string>();

            // Use reflexão para obter propriedades da classe
            PropertyInfo[] propriedades = objeto.GetType().GetProperties();

            foreach (CampoTagInfo campoTag in tagsDosCampos)
            {
                // Encontre a propriedade correspondente no objeto
                PropertyInfo propriedade = propriedades.FirstOrDefault(p => p.Name.Equals(campoTag.Nome, StringComparison.OrdinalIgnoreCase));

                if (propriedade != null)
                {
                    // Realize as críticas com base na tag e nos dados em DAO
                    if (campoTag.Tag == "O")
                    {
                        // Verifique se o valor na propriedade correspondente em DAO é vazio
                        object valor = propriedade.GetValue(objeto);
                        if (valor == null || (valor is string && string.IsNullOrEmpty((string)valor)))
                        {
                            criticas.Add($"O campo {propriedade.Name} é obrigatório.");
                        }
                    }
                    // Adicione outras verificações de tag conforme necessário
                }
            }
            MarcarControlesComErro(criticas);
            return criticas;
        }

        public List<tb.CampoTagInfo> LerTagsDosCamposDeTexto()
        {
            List<tb.CampoTagInfo> tags = new List<tb.CampoTagInfo>();

            foreach (Control control in Controls)
            {
                if (control is TextBox textBox)
                {
                    // Verifique se a Tag está definida como uma propriedade do TextBox (você pode personalizar isso)
                    if (textBox.Tag != null)
                    {
                        tags.Add(new CampoTagInfo
                        {
                            Nome = textBox.Name.Substring(3),
                            Tag = textBox.Tag.ToString()
                        });
                    }
                }
            }

            return tags;
        }

        private void MarcarControlesComErro(List<string> criticas)
        {
            // Uma lista para manter o registro dos nomes dos campos com erros
            HashSet<string> camposComErro = new HashSet<string>();

            foreach (string critica in criticas)
            {
                // Assumindo que a crítica contém o nome do campo, exemplo: "O campo Nome é obrigatório."
                string nomeCampo = critica.Split(' ')[2]; // Isso precisa ser ajustado conforme o formato da sua crítica
                camposComErro.Add(nomeCampo);
            }

            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is TextBox textBox)
                {
                    // Extrair o nome do campo do nome do controle, assumindo que o nome do controle começa com "txt"
                    string nomeCampo = textBox.Name.Substring(3);

                    if (camposComErro.Contains(nomeCampo))
                    {
                        // Marcar como erro
                        textBox.BackColor = Color.LightPink; // Cor para indicar erro
                    }
                    else
                    {
                        // Resetar a aparência para o normal
                        textBox.BackColor = SystemColors.Window;
                    }
                }
                // Adicione mais lógica aqui se houver outros tipos de controles
            }
        }

        #endregion

    }
}
