using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using System.Text;
using System.Threading.Tasks;
using System.IO; // A BIBLIOTECA DE ENTRADA E SAIDA DE ARQUIVOS
using iTextSharp; //E A BIBLIOTECA ITEXTSHARP E SUAS EXTENÇÕES
using iTextSharp.text; //ESTENSAO 1 (TEXT)
using iTextSharp.text.pdf;//ESTENSAO 2 (PDF)
using Excel = Microsoft.Office.Interop.Excel;
using GestorContratosMei;


namespace LerPlanilhaExcel
{
    public static class LerExcel
    {
        public static string CaminhoExcelSalvar { get; set; }
        public static string CaminhoExcelSelecionado { get; set; }

        public static bool Executando = false;

        public static int LINHA { get; set; }

        public static string CaminhoExcelSistema
        {
            get
            {
                return @"C:\Projects\GestorContratosMei\GestorContratosMei\excel\MeuExcel.xlsx";
            }

        }

        public static List<string> _ListaDeColunas = new List<string>();
        public static void ExecutarLista()
        {
            var workbook = new XLWorkbook(CaminhoExcelSelecionado);
            var sheet = workbook.Worksheet(1);
            int coluns = sheet.ColumnsUsed().Count();
            _ListaDeColunas = ManipularArquivos.ColetaListaDeColunas(coluns);

            System.Windows.Forms.MessageBox.Show("Total de Colunas " + _ListaDeColunas.Count.ToString());
        }

        public static void Executar()
        {
            Executando = true;
            var workbook = new XLWorkbook(CaminhoExcelSelecionado);
            StreamWriter logCnpj = new StreamWriter(@"C:\Projects\GestorContratosMei\GestorContratosMei\log\logsCnpj.txt");
            int contador = 0;

            int praca = workbook.Worksheets.Count;
            //Console.WriteLine(praca);

            int cont = 1;

            for (int i = 1; i <= praca; i++)
            {
                var sheet = workbook.Worksheet(i);

                var linha = 2;

                while (true)
                {
                    string cnpj = sheet.Cell("C" + linha).Value.ToString();
                    string email = sheet.Cell("O" + linha).Value.ToString();

                    /*var email = sheet.Cell("O" + linha).Value.ToString();
                    var cnpj = sheet.Cell("C" + linha).Value.ToString();
                    var emailLider = sheet.Cell("P" + linha).Value.ToString();
                    var dataContrato = sheet.Cell("Q" + linha).Value.ToString();
                    var nome = sheet.Cell("B" + linha).Value.ToString();
                    var endereco = sheet.Cell("J" + linha).Value.ToString();
                    var cep = sheet.Cell("L" + linha).Value.ToString();*/

                    if (contador == 5) break;
                    if (string.IsNullOrWhiteSpace(cnpj) && (string.IsNullOrWhiteSpace(email)))
                    {
                        contador++;
                    }
                    else
                    {
                        //Console.WriteLine(linha + ": " + sheet.ToString() + " - " + cnpj + " - " + email + " - " + emailLider + " - " + dataContrato);

                        logCnpj.Write(linha + " - " + sheet + " - " + cnpj + " - " + email + " - " + "\r\n");



                        DadosExcelModel dadosExcelModeL = new DadosExcelModel(sheet, linha);

                        ManipularArquivos.AtualizarPlanilha(dadosExcelModeL, cont);



                        cont++;
                    }

                    //System.Threading.Thread.Sleep(100);
                    linha++;

                }
                contador = 0;
            }

            workbook.Dispose();
            logCnpj.Close();
            //Console.WriteLine("Total de Praças lidas: " + praca);
            System.Windows.Forms.MessageBox.Show("Total de Praças lidas: " + praca);
            Executando = false;


        }



    }

}





