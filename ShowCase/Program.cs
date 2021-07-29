using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using ObjectToExcel;
using OfficeOpenXml;

namespace ShowCase
{
    class Car
    {
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
            cars.Add(new Car("GOLF 7 R", "HH101SD"));
            cars.Add(new Car("GOLF 7 R", "HH101SD"));
            cars.Add(new Car("GOLF 7 R", "HH101SD"));
            cars.Add(new Car("GOLF 7 R", "HH101SD"));
            cars.Add(new Car("GOLF 7 R", "HH101SD"));
            cars.Add(new Car("GOLF 7 R", "HH101SD"));
            cars.Add(new Car("GOLF 7 R", "HH101SD"));
            cars.Add(new Car("GOLF 7 R", "HH101SD"));


            string[] words = { "Alphabet", "Zebra", "ABC", "Αθήνα", "Москва" };

            string folder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage())
            {
                cars.ConvertToExcel(package, true);
                FileInfo fi = new FileInfo($"{folder}/cars.xlsx");
                // words.ConvertToExcel(package);
                // FileInfo fi = new FileInfo($"{folder}/words.xlsx");
                package.SaveAs(fi);

            }

            Console.WriteLine($"The List has been written");
        }
    }
}
