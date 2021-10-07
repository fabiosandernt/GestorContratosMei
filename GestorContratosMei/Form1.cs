using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using LerPlanilhaExcel;

namespace GestorContratosMei
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void buttonOpenExcel_Click(object sender, EventArgs e)
        {
            
            textBoxExcel.Text = ManipularArquivos.SelecionarArquivo();
                     

        }

        private void button1_Click(object sender, EventArgs e) //Executar
        {

            if (!string.IsNullOrEmpty(textBoxExcel.Text))
            {
                LerExcel.CaminhoExcelSelecionado = textBoxExcel.Text;

                ManipularArquivos.CriarArquivoSistema();

                buttonOpenExcel.Enabled = false;
                buttonExecutar.Enabled = false;
                //LerExcel.Executar();
                LerExcel.ExecutarLista();

                buttonOpenExcel.Enabled = true;
                buttonExecutar.Enabled = true;


            }
            else
            {

                MessageBox.Show("Escolha um arquivo", "Erro");
            }


            
            if (LerExcel._ListaDeColunas.Count > 0)
            {

                for (int i = 0; i < LerExcel._ListaDeColunas.Count; i++)
                {

                    string coluna = LerExcel._ListaDeColunas[i].ToString();
                    checkedListBoxColunas.Items.Add( coluna);

                }
            }
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBoxExcel_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
