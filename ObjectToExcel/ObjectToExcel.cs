using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace ObjectToExcel
{
    public static class ObjectToExcel
    {
        //
        // Summary:
        //    Converts a list of object type to an excel file.
        //
        // Parameters:
        //   items:
        //     The items to converted to an excel file.
        //   fileName:
        //     The name of the excel file.
        //   savePath:
        //     The path  where the excel file is to be saved.
        //   sheetName:
        //     The name of the sheet where the items is to be saved, default is items.
        //   fill:
        //     The value to be filled in if obejct property is null.
        //
        // Returns:
        //     The file path where the excel file was saved.
        //
        // Remarks:
        //     This is only public and still present to preserve compatibility with the V1 framework.
        public static async Task<string> ConvertToExcelAsync<T>(this IEnumerable<T> items, string fileName, string savePath, string sheetName = "Items", string fill = "null")
        {
            //remvoe nulls from items
            items = items.Where(i => i != null).ToList();

            var isSimpleType = IsSimpleType(items.FirstOrDefault().GetType());
            int headerCount = 1;
            if (!Directory.Exists(savePath))
            {
                Directory.CreateDirectory(savePath);
            }
            var filePath = Path.Combine(savePath, fileName);

            // delete file if it exists
            System.IO.File.Delete(filePath);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet workSheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == sheetName);
                if (workSheet == null)
                {
                    package.Workbook.Worksheets.Add(sheetName);
                }
                workSheet = package.Workbook.Worksheets[sheetName];
                int row = 1;
                int column = 1;

                if (!isSimpleType)
                {
                    var headers = typeof(T).GetProperties().Select(f => f.Name).ToList();
                    headerCount = headers.Count;
                    // create headers 
                    foreach (var header in headers)
                    {
                        workSheet.Cells[row, column].Value = header;
                        column++;
                    }
                }
                // Go To next row
                row++;
                // populate data
                foreach (var item in items)
                {
                    column = 1;
                    switch (isSimpleType)
                    {
                        case true:
                            workSheet.Cells[row, column].Value = item;
                            break;
                        default:
                            if (item != null)
                            {
                                foreach (PropertyInfo prop in item.GetType().GetProperties())
                                {
                                    try
                                    {
                                        var type = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
                                        workSheet.Cells[row, column].Value = prop.GetValue(item, null).ToString();

                                    }
                                    catch
                                    {
                                        workSheet.Cells[row, column].Value = fill;
                                    }
                                    column++;
                                }
                            }
                            break;
                    }
                    row++;
                }

                //Defining the tables parameters
                ExcelRange rg = workSheet.Cells[1, 1, items.Count() + 1, headerCount];
                string tableName = isSimpleType ? "items" : typeof(T).Name.ToString();

                //Ading a table to a Range
                ExcelTable tab = workSheet.Tables.Add(rg, tableName);

                //Formating the table style

                tab.TableStyle = TableStyles.Light14;
                // save excel file
                await package.SaveAsync();

                return filePath;
            }
        }
        public static bool IsSimpleType(Type type)
        {
            return
                type.IsPrimitive ||
                new Type[] {
                    typeof(string),
                    typeof(decimal),
                    typeof(DateTime),
                    typeof(DateTimeOffset),
                    typeof(TimeSpan),
                    typeof(Guid)
                }.Contains(type) ||
                type.IsEnum ||
                Convert.GetTypeCode(type) != TypeCode.Object ||
                (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>) && IsSimpleType(type.GetGenericArguments()[0]))
                ;
        }
    }
}
