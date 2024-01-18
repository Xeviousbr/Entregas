﻿using System;
using BonifacioEntregas.tb;
using System.Windows.Forms;
using System.Reflection;

namespace BonifacioEntregas
{
    public partial class fCadEntregadores : FormBase
    {
        private tb.Entregador clienteEspecifico;
        public fCadEntregadores()
        {
            InitializeComponent();
            base.DAO = new dao.EntregadorDAO();
            clienteEspecifico = DAO.GetUltimo() as tb.Entregador;
            base.reg = DAO.GetUltimo() as tb.Entregador;
            base.Mostra();
            base.LerTagsDosCamposDeTexto();
        }

        private void cntrole1_Load(object sender, EventArgs e)
        {

        }

        private void Teclou(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                base.Cancela();
            } else
            {
                base.cntrole1.EmEdicao = true;
            }            
        }

        private void dtpValidadeCNH_ValueChanged(object sender, EventArgs e)
        {
            if (Mostrando == false)
            {
                base.cntrole1.EmEdicao = true;
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

        private void cntrole1_AcaoRealizada_1(object sender, AcaoEventArgs e)
        {
            base.cntrole1_AcaoRealizada(sender, e, clienteEspecifico);
        }

        private void txtTelefone_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsDigit(e.KeyChar) && !char.Equals(e.KeyChar, '-') && !char.Equals(e.KeyChar, '/') && !char.Equals(e.KeyChar, '(') && !char.Equals(e.KeyChar, ')'))
            {
                e.Handled = true;
            }
        }
    }

}
