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
            public string Name { get; set; }
            public string Level { get; set; }
            public int Year { get; set; }
            public double FinalMark { get; set; }
        }
        [TestMethod]
        public async System.Threading.Tasks.Task TestPrimativeTypeList()
        {
            string[] words = { "Alphabet", "Zebra", "ABC", "Αθήνα", "Москва" };

            var file = await words.ConvertToExcelAsync("TestPrimativeTypeList.xls", Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));
            Assert.IsTrue(System.IO.File.Exists(file));
        }
        [TestMethod]
        public async System.Threading.Tasks.Task TestPrimativeTypeListWithNulls()
        {
            string?[] words = { "Alphabet", null, "Zebra", null, "ABC", "Αθήνα", "Москва" };

            var file = await words.ConvertToExcelAsync("TestPrimativeTypeListWithNulls.xls", Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));

            Assert.IsTrue(System.IO.File.Exists(file));
        }
        [TestMethod]
        public async System.Threading.Tasks.Task TestObjectListWithNulls()
        {
            Student[] students = {
                new Student() { Name="Gift"},
                null,
                new Student() { Name="Hlobile", Level="Grade 7", Year=2, FinalMark=80.00},
                null,
                new Student(){ FinalMark=3.0  },
                null
           };

            var file = await students.ConvertToExcelAsync("TestObjectListWithNulls.xls", Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));

            Assert.IsTrue(System.IO.File.Exists(file));
        }
    }
}
