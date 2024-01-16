using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;

namespace BonifacioEntregas
{
    public partial class Form2 : Form
    {
        private dao.EntregadorDAO entregadorDAO;
        private tb.Entregador reg;
        private int Direcao = 0;
        private bool EmAdicao = false;
        private bool Mostrando = false;

        public Form2()
        {
            InitializeComponent();
            entregadorDAO = new dao.EntregadorDAO();
            reg = entregadorDAO.GetUltimoEntregador();
            Mostra(); 
        }

        private void cntrole1_Load(object sender, EventArgs e)
        {

        }

        private void cntrole1_AcaoRealizada(object sender, AcaoEventArgs e)
        {
            switch (e.Acao)
            {                
                case "Adicionar":
                    LimparCampos();
                    EmAdicao = true;
                    break;
                case "Delete":
                    reg = entregadorDAO.Apagar(Direcao);
                    if (!Mostra())
                    {
                        if (Direcao==1)
                        {
                            cntrole1.Ultimo = true;
                        } else
                        {
                            cntrole1.Primeiro = true;
                        }                        
                    }
                    break;                    
                case "ParaTras":
                    Direcao = -1;
                    reg = entregadorDAO.ParaTraz();
                    if (!Mostra())
                    {
                        cntrole1.Ultimo = true;
                    }
                    break;
                case "ParaFrente":
                    Direcao = 1; ;
                    reg = entregadorDAO.ParaFrente();
                    if (!Mostra())
                    {
                        cntrole1.Primeiro = true;
                    }
                    break;
                case "Editar":
                    this.Text = "clicou";
                    break;
                case "CANC":
                    reg = entregadorDAO.GetEsse();
                    Mostra();
                    break; 
                case "OK":
                    Grava();
                    break; 
            }
        }

        private void Grava()
        {
            MapearCamposParaModelo(reg);
            if (EmAdicao)
            {
                reg.Id = 0;
            }
            entregadorDAO.Grava(reg);
            EmAdicao = false;
        }

        private void Teclou(object sender, KeyEventArgs e)
        {
            cntrole1.EmEdicao = true; 
        }

        private void LimparCampos()
        {
            foreach (Control control in this.Controls)
            {
                if (control is TextBox)
                {
                    control.Text = string.Empty;
                }
            }
        }

        #region Campos

        private bool Mostra()
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

        private void ProcessarTextBox(TextBox textBox)
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
            string propertyName = dtpControl.Name.Substring(3); // Remove o prefixo 'dtp'
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

        private void MapearCamposParaModelo(tb.Entregador reg)
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

        private void MapearTextBoxParaModelo(TextBox textBox, tb.Entregador reg)
        {
            string propertyName = textBox.Name.Substring(3); // Remove o prefixo 'txt'
            PropertyInfo propertyInfo = reg.GetType().GetProperty(propertyName);
            if (propertyInfo != null)
            {
                propertyInfo.SetValue(reg, textBox.Text, null);
            }
        }

        private void MapearDateTimePickerParaModelo(DateTimePicker dtpControl, tb.Entregador reg)
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

        private void dtpValidadeCNH_ValueChanged(object sender, EventArgs e)
        {
            if (Mostrando == false)
            {
                cntrole1.EmEdicao = true;
                DateTimePicker picker = sender as DateTimePicker;
                if (picker != null)
                {
                    string propertyName = picker.Name.Substring(3); // Remove o prefixo 'dtp'
                    PropertyInfo propertyInfo = reg.GetType().GetProperty(propertyName);
                    if (propertyInfo != null && (propertyInfo.PropertyType == typeof(DateTime) || propertyInfo.PropertyType == typeof(DateTime?)))
                    {
                        if (picker.Value != DateTime.MinValue)
                        {
                            picker.Format = DateTimePickerFormat.Short;
                            propertyInfo.SetValue(reg, picker.Value, null);
                        }
                        else
                        {
                            picker.CustomFormat = " ";
                            picker.Format = DateTimePickerFormat.Custom;
                            propertyInfo.SetValue(reg, null, null);
                        }
                    }
                }
            }
        }

    }

}
