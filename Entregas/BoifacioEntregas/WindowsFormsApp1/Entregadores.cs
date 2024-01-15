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

        public Form2()
        {
            InitializeComponent();
            entregadorDAO = new dao.EntregadorDAO();
            reg = entregadorDAO.GetUltimoEntregador();
            // cntrole1.Ultimo = true;
            Mostra(); 
        }

        private bool Mostra()
        {
            if (reg == null)
            {
                return false;
            }
            else
            {
                foreach (Control control in this.Controls)
                {
                    if (control is TextBox && control.Tag != null)
                    {
                        PropertyInfo propertyInfo = reg.GetType().GetProperty(control.Tag.ToString());
                        if (propertyInfo != null)
                        {
                            string valor = propertyInfo.GetValue(reg, null)?.ToString() ?? string.Empty;
                            control.Text = valor;
                        }
                    }
                }
                return true;
            }
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
        private void Apagar()
        {
            
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

        private void MapearCamposParaModelo(tb.Entregador reg)
        {
            foreach (Control control in this.Controls)
            {
                if (control is TextBox && control.Tag != null)
                {
                    try
                    {
                        PropertyInfo propertyInfo = reg.GetType().GetProperty(control.Tag.ToString());
                        if (propertyInfo != null)
                        {
                            propertyInfo.SetValue(reg, control.Text, null);
                        }
                    }
                    catch (Exception ex)
                    {
                        string x = ex.ToString();
                    }

                }
            }
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

    }

}
