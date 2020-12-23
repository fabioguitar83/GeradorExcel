using System;
using Xunit;
using ExcelGenerator.App;

namespace ExcelGenerator.Application.Test
{
    public class UnitTest1
    {
        [Fact]
        public void Test1()
        {

            var excelGenerator = new ExcelGenerator.App.ExcelGenerator();

            excelGenerator.CreateExcel();

        }
    }
}
