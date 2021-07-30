using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using Xunit.Sdk;
using ObjectToExcel;
using ShowCase;
using System.IO;
using System.Reflection;
using OfficeOpenXml;
using System.Linq;
using System;

namespace ObjectToExcelUnitTests
{
    [TestClass]
    public class UnitTest1
    {

        public class Student
        {
            [ExportToExcel(3)]
            public string Name { get; set; }
            public string Level { get; set; }
            [ExportToExcel(2)]
            public int Year { get; set; }
            [ExportToExcel(1)]
            public double FinalMark { get; set; }

            public override string ToString()
            {
                return $"Name: {Name}, Level: {Level}, Year: {Year}, FinalMark: {FinalMark}";
            }
        }
        [TestMethod]
        public void TestPrimativeTypeList()
        {
            IEnumerable<string> words = new string[] { "Alphabet", "Zebra", "ABC", "Αθήνα", "Москва" };

            string folder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
            string filePath = $"{folder}/TestPrimativeTypeList.xlsx";
            FileInfo fi = new FileInfo(filePath);
            string sheetName = "words";
            using (ExcelPackage package = new ExcelPackage())
            {
                words.ConvertToExcel(package, true, "words");
                package.SaveAs(fi);
                Assert.IsTrue(fi.Exists);
            }
            IEnumerable<string> newWords = new string[] { };
            using (ExcelPackage package = new ExcelPackage(fi))
            {

                newWords = newWords.LoadFromExcel(package, sheetName);
                Assert.IsTrue(!newWords.Except(words).Union(newWords.Except(words)).Any());
                // words.ConvertToExcel(package);
                // FileInfo fi = new FileInfo($"{folder}/words.xlsx");
                // package.SaveAs(fi);

            }
            System.IO.File.Delete(filePath);

        }
        [TestMethod]
        public void TestPrimativeTypeListWithNulls()
        {
            IEnumerable<string> words = new string[] { "Alphabet", "Zebra", "ABC", "Αθήνα", "Москва" };

            string folder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
            string filePath = $"{folder}/TestPrimativeTypeListWithNulls.xlsx";
            FileInfo fi = new FileInfo(filePath);
            string sheetName = "words";
            using (ExcelPackage package = new ExcelPackage())
            {
                words.ConvertToExcel(package, true, "words");
                package.SaveAs(fi);
                Assert.IsTrue(fi.Exists);
            }
            IEnumerable<string> newWords = new string[] { };
            using (ExcelPackage package = new ExcelPackage(fi))
            {

                newWords = newWords.LoadFromExcel(package, sheetName);
                Assert.IsTrue(!newWords.Except(words).Union(newWords.Except(words)).Any());
                // words.ConvertToExcel(package);
                // FileInfo fi = new FileInfo($"{folder}/words.xlsx");
                // package.SaveAs(fi);

            }
            System.IO.File.Delete(filePath);
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
            string filePath = $"{folder}/TestObjectListWithNulls.xlsx";
            string sheetName = "Students";
            FileInfo fi = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(fi))
            {
                students.ConvertToExcel(package, false, sheetName);
                package.SaveAs(fi);
                Assert.IsTrue(File.Exists(filePath));
            }
            IEnumerable<Student> newStudents = new Student[] { };
            using (ExcelPackage package = new ExcelPackage(fi))
            {
                newStudents = newStudents.LoadFromExcel(package, sheetName);
                students = students
                    .Where(i => i != null)
                    .Select(s => new Student
                    {
                        Name = s.Name,
                        Year = s.Year,
                        FinalMark = s.FinalMark
                    })
                    .ToArray();
                var names = students.Select(i => i.ToString()).ToList();
                var newNames = newStudents.Select(i => i.ToString()).ToList();

                students.ToList().ForEach(s => System.Console.WriteLine(s));
                System.Console.WriteLine("New Students: ");
                newStudents.ToList().ForEach(s => System.Console.WriteLine(s));

                var un = newNames.Except(names).Union(newNames.Except(names)).Any();
                Assert.IsTrue(!newNames.Except(names).Union(newNames.Except(names)).Any());

            }
            System.IO.File.Delete(filePath);
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
            string filePath = $"{folder}/TestObjectListWithNullsAll.xlsx";
            string sheetName = "Students";
            FileInfo fi = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(fi))
            {
                students.ConvertToExcel(package, true, sheetName);
                package.SaveAs(fi);
                Assert.IsTrue(File.Exists(filePath));
            }
            IEnumerable<Student> newStudents = new Student[] { };
            using (ExcelPackage package = new ExcelPackage(fi))
            {
                newStudents = newStudents.LoadFromExcel(package, sheetName);
                students = students.Where(i => i != null).ToArray();

                var names = students.Select(i => i.ToString()).ToList();
                var newNames = newStudents.Select(i => i.ToString()).ToList();
                // Assert.IsTrue(!File.Exists(filePath));


                var un = newNames.Except(names).Union(newNames.Except(names)).Any();
                Assert.IsTrue(!newNames.Except(names).Union(newNames.Except(names)).Any());
            }
            System.IO.File.Delete(filePath);

        }
    }
}
