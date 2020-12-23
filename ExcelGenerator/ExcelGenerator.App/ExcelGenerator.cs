using ClosedXML.Excel;
using ExcelGenerator.Model;
using System;

namespace ExcelGenerator.App
{
    public class ExcelGenerator : IExcelGenerator
    {
        public void CreateExcel()
        {
            using (var workbook = new XLWorkbook("C:\\Fabio\\testeGerador.xlsx"))
            {
                var worksheet = workbook.Worksheets.Worksheet("Cotacao");

                worksheet.Cell("A1:A3").Value = "Teste";

                workbook.Save();
            }
        }
    }
}
