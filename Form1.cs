using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;

namespace evaluacion_ceapsi
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            backgroundWorker1.WorkerReportsProgress = true;
            loadConfigList();
        }

        private bool fileAccepted = false;
        private ConfigurationInfo selectedConfigInfo;
        private List<ConfigurationInfo> configList;

        private void loadConfigList()
        {
            configList = ConfigurationInfo.listConfigurations();
            listBox1.DataSource = configList;
        }

        private bool fieldsAreValid()
        {
            int startingRow;
            int endRow;
            if (textBox1.Text.Trim().Equals(String.Empty))
            {
                MessageBox.Show("El archivo de datos es requerido");
                return false;
            }
            if (textBox4.Text.Trim().Equals(String.Empty))
            {
                MessageBox.Show("La carpeta de destino es requerida");
                return false;
            }
            if (listBox1.SelectedIndex == -1)
            {
                MessageBox.Show("Debe seleccionar una configuración");
                return false;
            }
            if (textBox2.Text.Trim().Equals(String.Empty))
            {
                MessageBox.Show("La fila inicial es requerida");
                return false;
            }
            else
            {
                if (!int.TryParse(textBox2.Text, out startingRow))
                {
                    MessageBox.Show("La fila inicial debe ser un número");
                    return false;
                }
            }
            if (textBox3.Text.Trim().Equals(String.Empty))
            {
                MessageBox.Show("La fila final es requerida");
                return false;
            }
            else
            {
                if (!int.TryParse(textBox3.Text, out endRow))
                {
                    MessageBox.Show("La fila final debe ser un número");
                    return false;
                }
            }
            if (endRow < startingRow)
            {
                MessageBox.Show("La fila final debe ser mayor o igual a la fila inicial");
                return false;
            }
            return true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (fieldsAreValid())
            {
                progressBar1.Value = 0;
                selectedConfigInfo = (ConfigurationInfo)listBox1.SelectedItem;
                backgroundWorker1.RunWorkerAsync();
                label6.Text = "Comenzando";
                button1.Enabled = false;
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            string sourceFile = textBox1.Text;
            string targetFile = selectedConfigInfo.TargetFile;
            string mappingFile = selectedConfigInfo.MappingFile;
            string resultDirectory = textBox4.Text;
            int startingRow = int.Parse(textBox2.Text.Trim());
            int endRow = int.Parse(textBox3.Text.Trim());
            MainProcessor.process(sourceFile, targetFile, mappingFile, resultDirectory, startingRow, endRow, backgroundWorker1);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            if (fileAccepted)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
            fileAccepted = false;
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            fileAccepted = true;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            textBox4.Text = folderBrowserDialog1.SelectedPath;
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            label6.Text = e.UserState.ToString();
            if (e.ProgressPercentage == 100)
            {
                button1.Enabled = true;
                label6.Text = "¡Listo!";
                MessageBox.Show("¡Listo!");
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            ConfigDialog configDialog = new ConfigDialog();
            DialogResult result = configDialog.ShowDialog();
            loadConfigList();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            selectedConfigInfo = (ConfigurationInfo)listBox1.SelectedItem;
            ConfigurationInfo.deleteConfiguration(selectedConfigInfo.Name);
            loadConfigList();
        }
    }
}
