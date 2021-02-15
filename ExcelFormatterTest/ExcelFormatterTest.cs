using NUnit.Framework;
using TempleScheduler;
using System;

namespace ExcelFormatterTest
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
            ExcelFormatter test = new ExcelFormatter("C:\\Users\\Home\\Desktop");
            test.FileNames();
        }
    }
}