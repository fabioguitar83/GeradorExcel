using NUnit.Framework;

namespace ExcelGenerator.App.Test
{
    public class Tests
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void Test1()
        {
            var excelGenerator = new ExcelGenerator();

            excelGenerator.CreateExcel();

            Assert.Pass();
        }
    }
}