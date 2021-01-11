using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace evaluacion_ceapsi
{
    public partial class ConfigDialog : Form
    {
        public ConfigDialog()
        {
            InitializeComponent();
        }

        private bool fileAccepted = false;

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Trim().Equals(String.Empty))
            {
                MessageBox.Show("El nombre es requerido");
                return;
            }
            if (textBox2.Text.Trim().Equals(String.Empty))
            {
                MessageBox.Show("El archivo plantilla es requerido");
                return;
            }
            if (textBox3.Text.Trim().Equals(String.Empty))
            {
                MessageBox.Show("El archivo de mapeo es requerido");
                return;
            }
            try
            {
                ConfigurationInfo.createNewConfiguration(textBox1.Text, textBox2.Text, textBox3.Text);
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            if (fileAccepted)
            {
                textBox2.Text = openFileDialog1.FileName;
            }
            fileAccepted = false;
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            fileAccepted = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            if (fileAccepted)
            {
                textBox3.Text = openFileDialog1.FileName;
            }
            fileAccepted = false;
        }
    }
}
