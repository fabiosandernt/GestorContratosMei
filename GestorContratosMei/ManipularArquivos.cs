using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.IO; // A BIBLIOTECA DE ENTRADA E SAIDA DE ARQUIVOS
using iTextSharp; //E A BIBLIOTECA ITEXTSHARP E SUAS EXTENÇÕES
using iTextSharp.text; //ESTENSAO 1 (TEXT)
using iTextSharp.text.pdf;//ESTENSAO 2 (PDF)
using LerPlanilhaExcel;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace GestorContratosMei
{

    public static class ManipularArquivos
    {
        public static IXLWorkbook wb;
        public static IXLWorksheet ws;
        public static string SelecionarArquivo()
        {
            OpenFileDialog openFDExcel = new OpenFileDialog();
            openFDExcel.Title = "Localize um Arquivo";
            openFDExcel.DefaultExt = ".xlsx";
            openFDExcel.FilterIndex = 0;
            openFDExcel.Filter = "xlsx|*.xlsx|xls|*.xls";
            openFDExcel.InitialDirectory = "";
            openFDExcel.Multiselect = false;

            if (openFDExcel.ShowDialog() == DialogResult.OK)
            {
                return openFDExcel.FileName;

            }
            return null;

        }
        public static void SalvarArquivos(string nomearquivo)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "Salvar Arquivo";
            saveFileDialog.FilterIndex = 0;
            saveFileDialog.Filter = "xlsx|*.xlsx|xls|*.xls";
            saveFileDialog.DefaultExt = ".xlsx";
            saveFileDialog.InitialDirectory = "";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                saveFileDialog.FileName = nomearquivo;

            }
        }
        public static void CriarArquivoSistema()
        {
            wb = new XLWorkbook();
            ws = wb.Worksheets.Add("Sheet");
            wb.SaveAs(LerExcel.CaminhoExcelSistema);

        }
        public static void AtualizarPlanilha(DadosExcelModel dadosExcelModel, int cont)
        {
            ws.Cell(cont, 1).Value = dadosExcelModel.Planilha;
            ws.Cell(cont, 2).Value = dadosExcelModel.Email;
            ws.Cell(cont, 3).Value = dadosExcelModel.Cnpj;
            ws.Cell(cont, 4).Value = dadosExcelModel.EmailLider;
            ws.Cell(cont, 5).Value = dadosExcelModel.DataContrato;
            ws.Cell(cont, 6).Value = dadosExcelModel.Nome;
            ws.Cell(cont, 7).Value = dadosExcelModel.Endereco;
            ws.Cell(cont, 8).Value = dadosExcelModel.Cep;
            wb.SaveAs(LerExcel.CaminhoExcelSistema);

        }

        public static List<string> _ListaDeColunas = new List<string>();
        
        //public static virtual void ColetaListaDeColunas(DadosExcelModel dadosExcelModel)
        //{

        //    if (_ListaDeColunas.Count == 0 )
        //    {
        //        //_ListaDeColunas.Add(dadosExcelModel.Email.ToString());
        //        //_ListaDeColunas.Add(dadosExcelModel.Cnpj.ToString());
        //        //_ListaDeColunas.Add(dadosExcelModel.EmailLider.ToString());
        //        //_ListaDeColunas.Add(dadosExcelModel.DataContrato.ToString());
        //        //_ListaDeColunas.Add(dadosExcelModel.Nome.ToString());
        //        //_ListaDeColunas.Add(dadosExcelModel.Endereco.ToString());
        //        //_ListaDeColunas.Add(dadosExcelModel.Cep.ToString());

        //    }
            
        //}

        public static List<string> ColetaListaDeColunas(int coluns)
        {
            string nomeColuna = "";
                                  

            for (int i = 0; i < coluns; i++)
            {
                nomeColuna = ws.Cell(1, i+1).Value.ToString();
                _ListaDeColunas.Add(nomeColuna);
            }
            
            return _ListaDeColunas;
        }
    }
}
