using System;
using System.Collections.Generic;
using ObjectToExcel;
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

        static async System.Threading.Tasks.Task Main(string[] args)
        {
            //List<Car> cars = new List<Car>();
            //cars.Add(new Car("GOLF 7 R", "HH101SD"));
            //cars.Add(new Car("GOLF 7 R", "HH101SD"));
            //cars.Add(new Car("GOLF 7 R", "HH101SD"));
            //cars.Add(new Car("GOLF 7 R", "HH101SD"));
            //cars.Add(new Car("GOLF 7 R", "HH101SD"));
            //cars.Add(new Car("GOLF 7 R", "HH101SD"));
            //cars.Add(new Car("GOLF 7 R", "HH101SD"));
            //cars.Add(new Car("GOLF 7 R", "HH101SD"));

            //var file = await cars.ConvertToExcelAsync("test.xls", "C:\\Users\\Giftm\\Downloads");

            string[] words = { "Alphabet", "Zebra", "ABC", "Αθήνα", "Москва" };

            var file = words.ConvertToExcel("test", "C:\\Users\\Giftm\\Downloads");

            Console.WriteLine($"The List has been written to: {file}");
        }
    }
}
