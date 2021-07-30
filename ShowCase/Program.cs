using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using ObjectToExcel;
using OfficeOpenXml;

namespace ShowCase
{
    class Car
    {
        public Car()
        {
        }

        public Car(string name, string reg)
        {
            Name = name;
            Registration = reg;
        }
        public string Name { get; set; }
        public string Registration { get; set; }
    }
    class Program
    {

        static void Main(string[] args)
        {
            List<Car> cars = new List<Car>();
            cars.Add(new Car("GOLF 6 R", "HH101SD"));
            cars.Add(null);
            cars.Add(new Car("GOLF 7 R", "HH101SD"));
            cars.Add(new Car("GOLF 8 R", "HH102SD"));
            cars.Add(null);
            cars.Add(new Car("GOLF Tiguan", "HH105SD"));


            string[] words = { "Alphabet", "Zebra", "ABC", "Αθήνα", "Москва" };

            string folder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = $"{folder}/cars.xlsx";
            System.IO.File.Delete(filePath);
            FileInfo fi = new FileInfo(filePath);
            string sheetName = "cars";
            using (ExcelPackage package = new ExcelPackage(fi))
            {
                cars.ConvertToExcel(package, true, sheetName);
                // words.ConvertToExcel(package);
                // FileInfo fi = new FileInfo($"{folder}/words.xlsx");
                package.SaveAs(fi);

            }

            Console.WriteLine($"The List has been written");
            List<Car> newCars = new List<Car>();
            IEnumerable<string> newWords = new string[] { };

            using (ExcelPackage package = new ExcelPackage(fi))
            {

                newCars = newCars.LoadFromExcel(package, sheetName).ToList();
                // words.ConvertToExcel(package);
                // FileInfo fi = new FileInfo($"{folder}/words.xlsx");
                // package.SaveAs(fi);

            }
            cars = cars.Where(o => o != null).ToList();
            // var un = newCars.Except(cars).Union(newCars.Except(cars)).Any();
            var firstNotSecond = newCars.Except(cars).ToList();
            var secondNotFirst = cars.Except(newCars).ToList();
            Console.WriteLine(!firstNotSecond.Any() && !secondNotFirst.Any());

        }
    }
}
