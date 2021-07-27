using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using Xunit.Sdk;
using ObjectToExcel;
using ShowCase;
using System.IO;
using System.Reflection;

namespace ObjectToExcelUnitTests
{
    [TestClass]
    public class UnitTest1
    {

        public class Student
        {
            [ExportToExcel(1)]
            public string Name { get; set; }
            public string Level { get; set; }
            [ExportToExcel(2)]
            public int Year { get; set; }
            public double FinalMark { get; set; }
        }
        [TestMethod]
        public void TestPrimativeTypeList()
        {
            string[] words = { "Alphabet", "Zebra", "ABC", "Αθήνα", "Москва" };

            var file = words.ConvertToExcel("TestPrimativeTypeList.xlsx", Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));
            Assert.IsTrue(System.IO.File.Exists(file));
        }
        [TestMethod]
        public void TestPrimativeTypeListWithNulls()
        {
            string?[] words = { "Alphabet", null, "Zebra", null, "ABC", "Αθήνα", "Москва" };

            var file = words.ConvertToExcel("TestPrimativeTypeListWithNulls.xlsx", Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));

            Assert.IsTrue(System.IO.File.Exists(file));
        }
        [TestMethod]
        public void TestObjectListWithNulls()
        {
            Student[] students = {
                new Student() { Name="Gift"},
                null,
                new Student() { Name="Hlobile", Level="Grade 7", Year=2, FinalMark=80.00},
                null,
                new Student(){ Name="Mkhosi",FinalMark=3.0  },
                null
           };

            var file = students.ConvertToExcel("TestObjectListWithNulls.xlsx", Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));

            Assert.IsTrue(System.IO.File.Exists(file));
        }
    }
}
