using NUnit.Framework;
using TempleScheduler;
using System;
using System.Threading.Tasks;

namespace ExcelFormatterTest
{
    public class Tests
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void TestingJSONDeserializer()
        {
            ExcelFormatter test = new ExcelFormatter(@"C:\Users\Home\Desktop", "");
            test.PersonsJSONDeserializer();
        }

        [Test]
        public async Task TestingExcelCreator()
        {
            ExcelFormatter test = new ExcelFormatter(@"C:\Users\Home\Desktop", "");
            await test.ExcelCreator();
        }
    }
}