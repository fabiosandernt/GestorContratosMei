using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace LerPlanilhaExcel
{
    public class DadosExcelModel
    {
        public string Email { get; set; }
        public string Cnpj { get; set; }
        public string EmailLider { get; set; }
        public string DataContrato { get; set; }
        public string Nome { get; set; }
        public string Endereco { get; set; }
        public string Cep { get; set; }
        public int Linha { get; set; }

        public string Planilha { get; set; }
        public DadosExcelModel(IXLWorksheet sheet, int linha)
        {
            Linha = linha;
            Planilha = sheet.Name.ToString();
            Email = sheet.Cell("O" + linha).Value.ToString();
            Cnpj = sheet.Cell("C" + linha).Value.ToString();
            EmailLider = sheet.Cell("P" + linha).Value.ToString();
            DataContrato = sheet.Cell("Q" + linha).Value.ToString();
            Nome = sheet.Cell("B" + linha).Value.ToString();
            Endereco = sheet.Cell("J" + linha).Value.ToString();
            Cep = sheet.Cell("L" + linha).Value.ToString();
            return;
        }
    }
}
