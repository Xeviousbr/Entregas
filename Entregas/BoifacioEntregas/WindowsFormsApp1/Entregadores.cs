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
    public partial class Form2 : FormBase
    {
        private dao.EntregadorDAO entregadorDAO;

        public Form2()
        {
            InitializeComponent();
            base.DAO = new dao.EntregadorDAO();
            base.reg = (tb.IDataEntity)DAO.GetUltimo();
            base.Mostra(); 
        }

        private void cntrole1_Load(object sender, EventArgs e)
        {

        }

        private void Teclou(object sender, KeyEventArgs e)
        {
            base.cntrole1.EmEdicao = true; 
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
            switch (e.Acao)
            {
                case "Adicionar":
                    base.LimparCampos();
                    base.EmAdicao = true;
                    break;
                case "Delete":
                    base.reg = (tb.Entregador)base.DAO.Apagar(Direcao);
                    if (!Mostra())
                    {
                        if (Direcao == 1)
                        {
                            base.cntrole1.Ultimo = true;
                        }
                        else
                        {
                            base.cntrole1.Primeiro = true;
                        }
                    }
                    break;
                case "ParaTras":
                    Direcao = -1;
                    base.reg = (tb.Entregador)base.DAO.ParaTraz();
                    if (!Mostra())
                    {
                        base.cntrole1.Ultimo = true;
                    }
                    break;
                case "ParaFrente":
                    Direcao = 1; ;
                    base.reg = (tb.Entregador)base.DAO.ParaFrente();
                    if (!Mostra())
                    {
                        base.cntrole1.Primeiro = true;
                    }
                    break;
                case "Editar":
                    this.Text = "clicou";
                    break;
                case "CANC":
                    base.reg = (tb.Entregador)base.DAO.GetEsse();
                    Mostra();
                    break;
                case "OK":
                    base.Grava();
                    break;
            }

        }
    }

}
