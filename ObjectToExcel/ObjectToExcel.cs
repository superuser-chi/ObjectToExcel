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
        public static ExcelPackage ConvertToExcel<T>(this IEnumerable<T> items, ExcelPackage package, string sheetName = "Items", string fill = "null")
        {
            //remvoe nulls from items
            items = items.Where(i => i != null).ToList();

            var isSimpleType = IsSimpleType(items.FirstOrDefault().GetType());
            int headerCount = 1;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
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
                headerCount = 0;
                // create headers 
                foreach (var header in headers)
                {
                    // add header when exported attribute exists
                    if (IsExported<T>(items.FirstOrDefault(), header))
                    {
                        workSheet.Cells[row, column].Value = header;
                        headerCount++;
                        column++;
                    }
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
                                if (IsExported<T>(items.FirstOrDefault(), prop.Name))
                                {
                                    try
                                    {
                                        var type = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
                                        ExportToExcel exporttribute = prop.GetCustomAttributes(typeof(ExportToExcel), true).Cast<ExportToExcel>().FirstOrDefault();
                                        workSheet.Cells[row, column].Value = prop.GetValue(item, null).ToString();
                                        column++;

                                    }
                                    catch
                                    {
                                        workSheet.Cells[row, column].Value = fill;
                                    }
                                }
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
            return package;

        }
        public static bool IsExported<T>(T item, string property)
        {
            var t = typeof(T);
            var pi = t.GetProperty(property);
            return Attribute.IsDefined(pi, typeof(ExportToExcel));
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
