﻿using BonifacioEntregas.tb;
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
        private List<CampoTagInfo> tagsDosCampos;

        public FormBase()
        {
            InitializeComponent();
            tagsDosCampos = new List<tb.CampoTagInfo>();
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

        protected void cntrole1_AcaoRealizada(object sender, AcaoEventArgs e, tb.IDataEntity entidade)
        {
            switch (e.Acao)
            {
                case "Adicionar":
                    LimparCampos();
                    EmAdicao = true;
                    break;
                case "Delete":
                    reg = DAO.Apagar(Direcao, entidade);
                    if (!Mostra())
                    {
                        if (Direcao == 1)
                        {
                            cntrole1.Ultimo = true;
                        }
                        else
                        {
                            cntrole1.Primeiro = true;
                        }
                    }
                    break;
                case "ParaTras":
                    Direcao = -1;
                    reg = DAO.ParaTraz();
                    if (!Mostra())
                    {
                        cntrole1.Ultimo = true;
                    }
                    break;
                case "ParaFrente":
                    Direcao = 1; ;
                    reg = DAO.ParaFrente();
                    if (!Mostra())
                    {
                        cntrole1.Primeiro = true;
                    }
                    break;
                case "Editar":
                    // this.Text = "clicou";
                    break;
                case "CANC":
                    Cancela();
                    break;
                case "OK":
                    Grava();
                    break;
            }
        }

        protected void Cancela()
        {
            reg = DAO.GetEsse();
            ResetarAparenciaControles();
            Mostra();
        }

        #endregion

        #region Crud
        protected void Grava()
        {
            MapearCamposParaModelo(DAO);
            List<string> criticas = FazerCriticas(DAO);            
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

        public List<string> FazerCriticas<T>(T objeto) where T : class
        {
            List<string> criticas = new List<string>();
            PropertyInfo[] propriedades = objeto.GetType().GetProperties();
            foreach (CampoTagInfo campoTag in tagsDosCampos)
            {
                PropertyInfo propriedade = propriedades.FirstOrDefault(p => p.Name.Equals(campoTag.Nome, StringComparison.OrdinalIgnoreCase));
                if (propriedade != null)
                {
                    if (campoTag.Tag == "O")
                    {
                        object valor = propriedade.GetValue(objeto);
                        if (valor == null || (valor is string && string.IsNullOrEmpty((string)valor)))
                        {
                            criticas.Add($"O campo {propriedade.Name} é obrigatório.");
                        }
                    }
                    else if (campoTag.Tag == "H" && propriedade.PropertyType == typeof(DateTime))
                    {
                        DateTime dataValor = (DateTime)propriedade.GetValue(objeto);
                        if (dataValor > DateTime.Today)
                        {
                            criticas.Add($"A data no campo {propriedade.Name} não pode ser posterior a hoje.");
                        }
                    }
                }
            }
            MarcarControlesComErro(criticas);
            return criticas;
        }

        public void LerTagsDosCamposDeTexto()
        {
            tagsDosCampos.Clear();
            foreach (Control control in Controls)
            {
                if (control is TextBox textBox)
                {
                    if (textBox.Tag != null)
                    {
                        tagsDosCampos.Add(new CampoTagInfo
                        {
                            Nome = textBox.Name.Substring(3),
                            Tag = textBox.Tag.ToString()
                        });
                    }
                }
                else if (control is DateTimePicker dtp)
                {
                    if (dtp.Tag != null)
                    {
                        tagsDosCampos.Add(new CampoTagInfo
                        {
                            Nome = dtp.Name.Substring(3),
                            Tag = dtp.Tag.ToString()
                        });
                    }
                }
            }
        }

        private void MarcarControlesComErro(List<string> criticas)
        {
            HashSet<string> camposComErro = new HashSet<string>();
            foreach (string critica in criticas)
            {
                string[] palavras = critica.Split(' ');
                int indiceCampo = Array.IndexOf(palavras, "campo");
                if (indiceCampo != -1 && indiceCampo + 1 < palavras.Length)
                {
                    string nomeCampo = palavras[indiceCampo + 1];
                    camposComErro.Add(nomeCampo);
                }
            }
            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is TextBox textBox)
                {
                    string nomeCampo = textBox.Name.Substring(3);
                    if (camposComErro.Contains(nomeCampo))
                    {
                        textBox.BackColor = Color.LightPink;
                    }
                    else
                    {
                        textBox.BackColor = SystemColors.Window;
                    }
                }
                else if (ctrl is DateTimePicker dtp)
                {
                    string nomeCampo = dtp.Name.Substring(3);
                    if (camposComErro.Contains(nomeCampo))
                    {
                        dtp.Font = new Font(dtp.Font.FontFamily, dtp.Font.Size, FontStyle.Bold);
                    }
                    else
                    {
                        dtp.Font = new Font(dtp.Font.FontFamily, dtp.Font.Size, FontStyle.Regular);
                    }
                }
            }
        }

        #endregion

    }
}
