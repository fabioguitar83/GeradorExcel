using ClosedXML.Excel;
using System;

namespace ExcelGenerator.Console
{
    class ExcelGenerator
    {

        public void CreateExcel()
        {

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Detalhes da Cotação");

                GeneralSettings(worksheet);
                AddHeaders(worksheet, DateTime.Now);
                AddRQF(worksheet);
                AddRemuneracao(worksheet);
                AddFluxoCaixa(worksheet);
                AddEventos(worksheet);
                AddPrecificacao(worksheet);

                workbook.SaveAs("c://Fabio/teste.xlsx");
            }
        }

        private void GeneralSettings(IXLWorksheet worksheet)
        {
            worksheet.Style.Font.FontName = "Calibri";
            worksheet.Style.Font.FontSize = 11;

            worksheet.Column("A").Width = 29;
            worksheet.Column("B").Width = 29;
            worksheet.Column("C").Width = 29;
            worksheet.Column("D").Width = 29;
            worksheet.Column("E").Width = 29;

            worksheet.Columns(1,100).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            worksheet.Columns(1,100).Style.Border.LeftBorderColor = XLColor.White;

            worksheet.Columns(1, 100).Style.Border.RightBorder = XLBorderStyleValues.Thin;
            worksheet.Columns(1, 100).Style.Border.RightBorderColor = XLColor.White;

            worksheet.Columns(1, 100).Style.Border.TopBorder = XLBorderStyleValues.Thin;
            worksheet.Columns(1, 100).Style.Border.TopBorderColor = XLColor.White;

            worksheet.Columns(1, 100).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            worksheet.Columns(1, 100).Style.Border.BottomBorderColor = XLColor.White;

        }

        private void AddHeaders(IXLWorksheet worksheet, DateTime dataCotacao)
        {
            worksheet.Cell("A1").Value = "Detalhe da cotação";
            worksheet.Cell("A1").Style.Font.FontSize = 18;
            worksheet.Cell("A1").Style.Font.Bold = true;


            worksheet.Cell("C3").Value = "Data Cotação";
            worksheet.Cell("C3").Style.Font.Bold = true;
            worksheet.Cell("C3").Style.Alignment.Indent = 1;

            worksheet.Cell("D3").Value = dataCotacao.ToString("dd/MM/yyyy");
            worksheet.Cell("D3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
        }

        private void AddRQF(IXLWorksheet worksheet)
        {
            //header
            worksheet.Cell("A4").Value = "RQF";
            worksheet.Cell("A4").Style.Font.Bold = true;
            worksheet.Cell("A4").Style.Font.FontSize = 13.5;

            //bloco da esquerda
            worksheet.Cell("A6").Value = "Apelido";
            worksheet.Cell("B6").Value = "teste";

            worksheet.Cell("A7").Value = "Data inicio janela";
            worksheet.Cell("B7").Value = "teste";

            worksheet.Cell("A8").Value = "Commitment Fee";
            worksheet.Cell("B8").Value = "teste";

            worksheet.Cell("A9").Value = "Tipo Operação";
            worksheet.Cell("B9").Value = "teste";

            worksheet.Cell("A10").Value = "Valor Financeiro";
            worksheet.Cell("B10").Value = "teste";

            worksheet.Cell("A11").Value = "Prazo";
            worksheet.Cell("B11").Value = "teste";

            worksheet.Cell("A12").Value = "Data vencimento";
            worksheet.Cell("B12").Value = "teste";

            //bloco da direita
            worksheet.Cell("C6").Value = "Desembolso";
            worksheet.Cell("D6").Value = "teste";

            worksheet.Cell("C7").Value = "Data fim janela";
            worksheet.Cell("D7").Value = "teste";

            worksheet.Cell("C8").Value = "Fee de Arrependimento";
            worksheet.Cell("D8").Value = "teste";

            worksheet.Cell("C9").Value = "Moeda";
            worksheet.Cell("D9").Value = "teste";

            worksheet.Cell("C10").Value = "Frequência Prazo";
            worksheet.Cell("D10").Value = "teste";

            worksheet.Cell("C11").Value = "Data desembolso";
            worksheet.Cell("D11").Value = "teste";

            var rangeBoldEsquerda = worksheet.Range("A6:A12");
            var rangeBoldDireita = worksheet.Range("C6:C11");

            rangeBoldEsquerda.Style.Alignment.Indent = 1;
            rangeBoldEsquerda.Style.Font.Bold = true;

            rangeBoldDireita.Style.Alignment.Indent = 1;
            rangeBoldDireita.Style.Font.Bold = true;

            var rangeBorderDireita = worksheet.Range("B6:B12");

            rangeBorderDireita.Style.Border.RightBorderColor = XLColor.Black;
            rangeBorderDireita.Style.Border.RightBorder = XLBorderStyleValues.Thin;
        }

        private void AddRemuneracao(IXLWorksheet worksheet)
        {
            //header
            worksheet.Cell("A14").Value = "Remuneração";
            worksheet.Cell("A14").Style.Font.Bold = true;
            worksheet.Cell("A14").Style.Font.FontSize = 13.5;

            //bloco da esquerda
            worksheet.Cell("A16").Value = "Indexadores";
            worksheet.Cell("B16").Value = "teste";

            worksheet.Cell("A17").Value = "Regime de juros";
            worksheet.Cell("B17").Value = "teste";

            worksheet.Cell("A18").Value = "Modo de correção";
            worksheet.Cell("B18").Value = "teste";

            worksheet.Cell("A19").Value = "Dia do aniversário";
            worksheet.Cell("B19").Value = "teste";

            worksheet.Cell("A20").Value = "Mês ineficiência";
            worksheet.Cell("B20").Value = "teste";

            worksheet.Cell("A21").Value = "Condição de Resgate";
            worksheet.Cell("B21").Value = "teste";

            //bloco da direita
            worksheet.Cell("C16").Value = "Método de juros";
            worksheet.Cell("D16").Value = "teste";

            worksheet.Cell("C17").Value = "Contagem dias";
            worksheet.Cell("D17").Value = "teste";

            worksheet.Cell("C18").Value = "Fixing";
            worksheet.Cell("D18").Value = "teste";

            worksheet.Cell("C19").Value = "Ineficiência";
            worksheet.Cell("D19").Value = "teste";

            worksheet.Cell("C20").Value = "Critério pro-rata";
            worksheet.Cell("D20").Value = "teste";

            var rangeBoldEsquerda = worksheet.Range("A16:A21");
            var rangeBoldDireita = worksheet.Range("C16:C21");

            rangeBoldEsquerda.Style.Alignment.Indent = 1;
            rangeBoldEsquerda.Style.Font.Bold = true;

            rangeBoldDireita.Style.Alignment.Indent = 1;
            rangeBoldDireita.Style.Font.Bold = true;

            var rangeBorderDireita = worksheet.Range("B16:B21");

            rangeBorderDireita.Style.Border.RightBorderColor = XLColor.Black;
            rangeBorderDireita.Style.Border.RightBorder = XLBorderStyleValues.Thin;
        }

        private void AddFluxoCaixa(IXLWorksheet worksheet)
        {
            //header
            worksheet.Cell("A23").Value = "Fluxo de Caixa";
            worksheet.Cell("A23").Style.Font.Bold = true;
            worksheet.Cell("A23").Style.Font.FontSize = 13.5;

            //bloco da esquerda
            worksheet.Cell("A25").Value = "Modalidade juros";
            worksheet.Cell("B25").Value = "teste";

            worksheet.Cell("A26").Value = "Periodicidade juros";
            worksheet.Cell("B26").Value = "teste";

            worksheet.Cell("A27").Value = "Data incorpora juros";
            worksheet.Cell("B27").Value = "teste";

            worksheet.Cell("A28").Value = "Modalidade amortização";
            worksheet.Cell("B28").Value = "teste";

            worksheet.Cell("A29").Value = "Periodicidade amortização";
            worksheet.Cell("B29").Value = "teste";

            //bloco da direita
            worksheet.Cell("C25").Value = "Frequência juros";
            worksheet.Cell("D25").Value = "teste";

            worksheet.Cell("C26").Value = "Incorpora juros";
            worksheet.Cell("D26").Value = "teste";

            worksheet.Cell("C27").Value = "1º pagamento de Juros";
            worksheet.Cell("D27").Value = "teste";

            worksheet.Cell("C28").Value = "Frequência amortização";
            worksheet.Cell("D28").Value = "teste";

            worksheet.Cell("C29").Value = "1º pagamento de Amortização";
            worksheet.Cell("D29").Value = "teste";

            var rangeBoldEsquerda = worksheet.Range("A25:A29");
            var rangeBoldDireita = worksheet.Range("C25:C29");

            rangeBoldEsquerda.Style.Alignment.Indent = 1;
            rangeBoldEsquerda.Style.Font.Bold = true;

            rangeBoldDireita.Style.Alignment.Indent = 1;
            rangeBoldDireita.Style.Font.Bold = true;

            var rangeBorderDireita = worksheet.Range("B25:B29");

            rangeBorderDireita.Style.Border.RightBorderColor = XLColor.Black;
            rangeBorderDireita.Style.Border.RightBorder = XLBorderStyleValues.Thin;
        }

        private void AddEventos(IXLWorksheet worksheet)
        {
            //header
            worksheet.Cell("A31").Value = "Evento";
            worksheet.Cell("B31").Value = "Data Início Ajustada";
            worksheet.Cell("C31").Value = "Data Liquidação";
            worksheet.Cell("D31").Value = "% Amt.";
            worksheet.Cell("E31").Value = "Valor Parcela";

            var rangeHeader = worksheet.Range("A31:E31");
            rangeHeader.Style.Font.Bold = true;

            //Tabela
            worksheet.Cell("A32").Value = "Desembolso";
            worksheet.Cell("B32").Value = "teste";
            worksheet.Cell("C32").Value = "teste";
            worksheet.Cell("D32").Value = "teste";
            worksheet.Cell("E32").Value = "teste";

            //FOR DE JUROS
            worksheet.Cell("A33").Value = "Juros";
            worksheet.Cell("B33").Value = "teste";
            worksheet.Cell("C33").Value = "teste";
            worksheet.Cell("D33").Value = "teste";
            worksheet.Cell("E33").Value = "teste";

            worksheet.Cell("A34").Value = "Juros";
            worksheet.Cell("B34").Value = "teste";
            worksheet.Cell("C34").Value = "teste";
            worksheet.Cell("D34").Value = "teste";
            worksheet.Cell("E34").Value = "teste";

            worksheet.Cell("A35").Value = "Juros";
            worksheet.Cell("B35").Value = "teste";
            worksheet.Cell("C35").Value = "teste";
            worksheet.Cell("D35").Value = "teste";
            worksheet.Cell("E35").Value = "teste";

            worksheet.Cell("A36").Value = "Juros";
            worksheet.Cell("B36").Value = "teste";
            worksheet.Cell("C36").Value = "teste";
            worksheet.Cell("D36").Value = "teste";
            worksheet.Cell("E36").Value = "teste";

            //FIM FOR DE JUROS

            worksheet.Cell("A37").Value = "Amortização";
            worksheet.Cell("B37").Value = "teste";
            worksheet.Cell("C37").Value = "teste";
            worksheet.Cell("D37").Value = "teste";
            worksheet.Cell("E37").Value = "teste";

            worksheet.Cell("A38").Value = "Vencimento";
            worksheet.Cell("B38").Value = "teste";
            worksheet.Cell("C38").Value = "teste";
            worksheet.Cell("D38").Value = "teste";
            worksheet.Cell("E38").Value = "teste";
        }

        private void AddPrecificacao(IXLWorksheet worksheet)
        {
            //header
            worksheet.Cell("A40").Value = "Precificação";
            worksheet.Cell("A40").Style.Font.Bold = true;
            worksheet.Cell("A40").Style.Font.FontSize = 13.5;

            //tabela            
            worksheet.Cell("A41").Value = "CDI+";

            worksheet.Cell("A42").Value = "Percentual Indexador";
            worksheet.Cell("B42").Value = "teste";

            worksheet.Cell("A43").Value = "Taxa Fixa";
            worksheet.Cell("B43").Value = "teste";

            worksheet.Cell("A44").Value = "Liquidação FU";
            worksheet.Cell("B44").Value = "teste";

            worksheet.Cell("A45").Value = "Mesa Parte";
            worksheet.Cell("B45").Value = "teste";

            worksheet.Cell("A46").Value = "Mesa Contraparte";
            worksheet.Cell("B46").Value = "teste";

            worksheet.Cell("A47").Value = "Estratégia Parte";
            worksheet.Cell("B47").Value = "teste";

            worksheet.Cell("A48").Value = "Estratégia Contraparte";
            worksheet.Cell("B48").Value = "teste";

        }
    }
}
