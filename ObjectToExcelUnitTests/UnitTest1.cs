using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using Xunit.Sdk;
using ObjectToExcel;
using ShowCase;
using System.IO;
using System.Reflection;
using OfficeOpenXml;

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
            [ExportToExcel(3)]
            public int Year { get; set; }
            [ExportToExcel(2)]
            public double FinalMark { get; set; }
        }
        [TestMethod]
        public void TestPrimativeTypeList()
        {
            string[] words = { "Alphabet", "Zebra", "ABC", "Αθήνα", "Москва" };

            string folder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
            using (ExcelPackage package = new ExcelPackage())
            {
                words.ConvertToExcel(package);
                FileInfo fi = new FileInfo($"{folder}/TestPrimativeTypeList.xlsx");
                package.SaveAs(fi);
                Assert.IsTrue(fi.Exists);
            }
        }
        [TestMethod]
        public void TestPrimativeTypeListWithNulls()
        {
            string?[] words = { "Alphabet", null, "Zebra", null, "ABC", "Αθήνα", "Москва" };

            string folder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
            using (ExcelPackage package = new ExcelPackage())
            {
                words.ConvertToExcel(package);
                FileInfo fi = new FileInfo($"{folder}/TestPrimativeTypeListWithNulls.xlsx");
                package.SaveAs(fi);
                Assert.IsTrue(fi.Exists);
            }
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

            string folder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
            using (ExcelPackage package = new ExcelPackage())
            {
                students.ConvertToExcel(package, false);
                FileInfo fi = new FileInfo($"{folder}/TestObjectListWithNulls.xlsx");
                package.SaveAs(fi);
                Assert.IsTrue(fi.Exists);
            }

        }
        [TestMethod]
        public void TestObjectListWithNullsAll()
        {
            Student[] students = {
                new Student() { Name="Gift"},
                null,
                new Student() { Name="Hlobile", Level="Grade 7", Year=2, FinalMark=80.00},
                null,
                new Student(){ Name="Mkhosi",FinalMark=3.0  },
                null
           };

            string folder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
            using (ExcelPackage package = new ExcelPackage())
            {
                students.ConvertToExcel(package);
                FileInfo fi = new FileInfo($"{folder}/TestObjectListWithNullsAll.xlsx");
                package.SaveAs(fi);
                Assert.IsTrue(fi.Exists);
            }

        }
    }
}
