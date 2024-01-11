using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            // 
        }

        private void AbrirOuFocarFormulario<T>() where T : Form, new()
        {
            // Verifica se já existe uma instância do formulário
            T formExistente = Application.OpenForms.OfType<T>().FirstOrDefault();

            if (formExistente != null)
            {
                // Se a janela estiver minimizada, restaura para o estado normal
                if (formExistente.WindowState == FormWindowState.Minimized)
                {
                    formExistente.WindowState = FormWindowState.Normal;
                }

                // Traz a janela para o primeiro plano
                formExistente.BringToFront();
                formExistente.Focus();
            }
            else
            {
                // Se não houver uma instância existente, cria uma nova
                T novoForm = new T();
                novoForm.Show();
            }
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            AbrirOuFocarFormulario<Form2>();
        }

        private void pictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            PictureBox pb = sender as PictureBox;
            if (pb != null)
            {
                pb.BorderStyle = BorderStyle.Fixed3D; // Efeito de pressionado
            }
        }

        private void pictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            PictureBox pb = sender as PictureBox;
            if (pb != null)
            {
                pb.BorderStyle = BorderStyle.FixedSingle; // Volta ao estilo original
            }
        }


    }

}
