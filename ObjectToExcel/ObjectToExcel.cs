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
        public static ExcelPackage ConvertToExcel<T>(this IEnumerable<T> items, ExcelPackage package, bool printAll = true, string sheetName = "Items", string fill = null)
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
                headerCount = 0;
                int order = 0;
                // create headers 
                List<Column> columns = new List<Column>();
                foreach (PropertyInfo prop in typeof(T).GetProperties())
                {
                    // add header when exported attribute exists
                    if (printAll || IsExported<T>(items.FirstOrDefault(), prop.Name))
                    {
                        ExportToExcel exporttribute = prop.GetCustomAttributes(typeof(ExportToExcel), true).Cast<ExportToExcel>().FirstOrDefault();
                        columns.Add(new Column
                        {
                            Order = GetOrder(exporttribute, printAll, order),
                            Name = prop.Name,
                            Value = prop.Name
                        });
                        headerCount++;
                        order++;
                    }
                }
                columns.Sort((x, y) => x.Order.CompareTo(y.Order));
                workSheet.Cells[row, column].LoadFromText(String.Join(",", columns.Select(x => x.Value).ToList()));

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

                            int order = 1;
                            List<Column> columns = new List<Column>();
                            foreach (PropertyInfo prop in item.GetType().GetProperties())
                            {
                                if (printAll || IsExported<T>(items.FirstOrDefault(), prop.Name))
                                {
                                    ExportToExcel exporttribute = prop.GetCustomAttributes(typeof(ExportToExcel), true).Cast<ExportToExcel>().FirstOrDefault();
                                    try
                                    {
                                        columns.Add(new Column
                                        {
                                            Order = GetOrder(exporttribute, printAll, order),
                                            Name = prop.Name,
                                            Value = prop.GetValue(item, null).ToString()
                                        });
                                        order++;
                                        // workSheet.Cells[row, column].Value = prop.GetValue(item, null).ToString();
                                    }
                                    catch
                                    {
                                        workSheet.Cells[row, column].Value = fill;
                                        columns.Add(new Column
                                        {
                                            Order = exporttribute == null ? order : exporttribute.order,
                                            Name = prop.Name,
                                            Value = fill
                                        });
                                    }
                                }
                            }
                            columns.Sort((x, y) => x.Order.CompareTo(y.Order));
                            workSheet.Cells[row, column].LoadFromText(String.Join(",", columns.Select(x => x.Value).ToList()));
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
        //
        // Summary:
        //    Loads an excel file into a  list.
        //
        // Parameters:
        //   collection:
        //     The collection in which the data is to be loaded into.
        //   package:
        //     The EPPLUS package to be used in reading in the data.
        //   savePath:
        //     The path  where the excel file is to be saved.
        //   sheetName:
        //     The name of the sheet where the items is to be saved, default is items.
        //   readAll:
        //     A boolean that describes if all columns of the object are to be read in.
        //
        // Returns:
        //     The file path where the excel file was saved.
        //
        // Remarks:
        //     This is only public and still present to preserve compatibility with the V1 framework.
        public static IEnumerable<T> LoadFromExcel<T>(this IEnumerable<T> collection, ExcelPackage package, string sheetName, bool readAll = true)
        {
            return IsSimpleType(typeof(T))
                ? collection.LoadPrimitiveFromExcel(package, sheetName, readAll)
                : collection.LoadObjectsFromExcel(package, sheetName, readAll);

        }
        public static IEnumerable<T> LoadPrimitiveFromExcel<T>(this IEnumerable<T> collection, ExcelPackage package, string sheetName, bool readAll = true)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == sheetName);

            var rows = worksheet.Cells
                .Select(cell => cell.Start.Row)
                .Distinct()
                .OrderBy(x => x);

            if (IsSimpleType(typeof(T)))
            {
                collection = rows.Skip(1)
                .Select(row =>
                {
                    var val = worksheet.Cells[row, 1];
                    T tnew = val.GetValue<T>();
                    return tnew;
                });
            }
            return collection;
        }

        public static IEnumerable<T> LoadObjectsFromExcel<T>(this IEnumerable<T> collection, ExcelPackage package, string sheetName, bool readAll = true)
        {

            Func<CustomAttributeData, bool> columnOnly = y => y.AttributeType == typeof(ExportToExcel);

            var columns = typeof(T)
                .GetProperties()
                .Where(x => readAll || x.CustomAttributes.Any(columnOnly))
            .Select((p, order) => new
            {
                Property = p,
                Column = GetOrder(p.GetCustomAttributes<ExportToExcel>().FirstOrDefault(), readAll, order + 1)//safe because if where above
            }).ToList();

            ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == sheetName);

            var rows = worksheet.Cells
                .Select(cell => cell.Start.Row)
                .Distinct()
                .OrderBy(x => x);

            //Create the collection container
            collection = rows.Skip(1)
                    .Select(row =>
                    {
                        T tnew = (T)Activator.CreateInstance(typeof(T));
                        columns.ForEach(col =>
                            {
                                //This is the real wrinkle to using reflection - Excel stores all numbers as double including int
                                var location = worksheet.GetHeaderColumns().FirstOrDefault(i => i.Value == col.Property.Name);
                                if (location.Key == 0)
                                {
                                    col.Property.SetValue(tnew, null);
                                    return;
                                }
                                // Console.WriteLine($"col.Property.Name: {col.Property.Name}");
                                // Console.WriteLine($"column: {location.Key}");
                                var val = worksheet.Cells[row, location.Key];
                                // col.Property.Name ==
                                //If it is numeric it is a double since that is how excel stores all numbers
                                if (val.Value == null)
                                {
                                    col.Property.SetValue(tnew, null);
                                    return;
                                }
                                if (col.Property.PropertyType == typeof(short))
                                {
                                    col.Property.SetValue(tnew, val.GetValue<short>());
                                    return;
                                }
                                if (col.Property.PropertyType == typeof(ushort))
                                {
                                    col.Property.SetValue(tnew, val.GetValue<ushort>());
                                    return;
                                }
                                if (col.Property.PropertyType == typeof(int))
                                {
                                    col.Property.SetValue(tnew, val.GetValue<int>());
                                    return;
                                }
                                if (col.Property.PropertyType == typeof(uint))
                                {
                                    col.Property.SetValue(tnew, val.GetValue<uint>());
                                    return;
                                }
                                if (col.Property.PropertyType == typeof(bool))
                                {
                                    col.Property.SetValue(tnew, val.GetValue<bool>());
                                    return;
                                }
                                if (col.Property.PropertyType == typeof(long))
                                {
                                    col.Property.SetValue(tnew, val.GetValue<long>());
                                    return;
                                }
                                if (col.Property.PropertyType == typeof(ulong))
                                {
                                    col.Property.SetValue(tnew, val.GetValue<ulong>());
                                    return;
                                }
                                if (col.Property.PropertyType == typeof(double))
                                {
                                    col.Property.SetValue(tnew, val.GetValue<double>());
                                    return;
                                }
                                if (col.Property.PropertyType == typeof(decimal))
                                {
                                    col.Property.SetValue(tnew, val.GetValue<decimal>());
                                    return;
                                }
                                if (col.Property.PropertyType == typeof(DateTime))
                                {
                                    col.Property.SetValue(tnew, val.GetValue<DateTime>());
                                    return;
                                }
                                //Its a string
                                col.Property.SetValue(tnew, val.GetValue<string>());
                            });

                        return tnew;
                    });


            //Send it back
            return collection;
        }

        public static Dictionary<int, string> GetHeaderColumns(this ExcelWorksheet sheet)
        {
            var columnNames = new Dictionary<int, string>();
            foreach (var firstRowCell in sheet.Cells[sheet.Dimension.Start.Row, sheet.Dimension.Start.Column, 1, sheet.Dimension.End.Column])
            {
                columnNames.Add(firstRowCell.Start.Column, firstRowCell.Text);
            }
            return columnNames;
        }
        private static int GetOrder(ExportToExcel exportExcel, bool all, int order)
        {
            if (all) return order;
            return exportExcel == null ? order : exportExcel.order;
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
