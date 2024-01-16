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

        #endregion

        #region Crud
        protected void Grava()
        {
            MapearCamposParaModelo(DAO);
            if (EmAdicao)
            {
                DAO.Adicao = true;
            }
            DAO.Grava(DAO);
            EmAdicao = false;
        }

        #endregion
    }
}
