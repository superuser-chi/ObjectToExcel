# Object-To-Excel

## Background

This is a simple library that can be used to convert a list to excel and read in data from an excel file.

## Examples

See the following program for examples

To read in a list in a list you can use the following method:

    ```csharp

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
            }

        }

    ```
