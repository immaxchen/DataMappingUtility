using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace DataMappingUtility
{
    public static class DataIO
    {
        public static List<List<string>> ReadDataString(string dataString, string delimiter = ",", string linebreak = "\r\n")
        {
            return dataString.Split(new string[] { linebreak }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Split(new string[] { delimiter }, StringSplitOptions.None).ToList())
                .ToList();
        }

        public static string ToDataString(this List<List<string>> dataTable, string delimiter = ",", string linebreak = "\r\n")
        {
            return String.Join(linebreak,
                dataTable.Select(x => String.Join(delimiter, x.ToArray())).ToArray()
            );
        }

        public static List<List<string>> ReadExcel(string filepath, string sheetname = null)
        {
            List<List<string>> dataTable = new List<List<string>>();
            using (ExcelPackage excel = new ExcelPackage(new FileInfo(filepath)))
            {
                if (sheetname == null) sheetname = excel.Workbook.Worksheets.First().Name;
                ExcelWorksheet sheet = excel.Workbook.Worksheets[sheetname];
                for (int i = sheet.Dimension.Start.Row; i <= sheet.Dimension.End.Row; i++)
                {
                    dataTable.Add(new List<string>());
                    var record = dataTable.Last();
                    for (int j = sheet.Dimension.Start.Column; j <= sheet.Dimension.End.Column; j++)
                        record.Add(sheet.Cells[i, j].Value.ToString());
                }
            }
            return dataTable;
        }

        public static void ToExcel(this List<List<string>> dataTable, string filepath, string sheetname = null)
        {
            using (ExcelPackage excel = new ExcelPackage(new FileInfo(filepath)))
            {
                if (sheetname == null && excel.Workbook.Worksheets.Count == 0) sheetname = "Sheet1";
                if (sheetname == null) sheetname = excel.Workbook.Worksheets.First().Name;
                if (excel.Workbook.Worksheets.Select(x => x.Name).Contains(sheetname)) excel.Workbook.Worksheets.Delete(sheetname);
                ExcelWorksheet sheet = excel.Workbook.Worksheets.Add(sheetname);
                for (int i = 0; i < dataTable.Count; i++)
                    for (int j = 0; j < dataTable[0].Count; j++)
                        sheet.Cells[i + 1, j + 1].Value = dataTable[i][j];
                excel.Save();
            }
        }
    }

    public static class DataOOP
    {
        public static IEnumerable<T> Generate<T>(this List<List<string>> dataTable, T dataModel)
        {
            if (dataTable.Count == 0) return null;
            var ctor = dataModel.GetType().GetConstructors().Single();
            var pars = ctor.GetParameters();
            List<T> dataObjects = new List<T>();
            List<string> header = dataTable[0];
            for (int i = 1; i < dataTable.Count; i++)
            {
                List<object> args = new List<object>();
                foreach (var p in pars)
                {
                    string val = null;
                    if (header.Contains(p.Name))
                        val = dataTable[i][header.IndexOf(p.Name)];
                    if (p.ParameterType == typeof(int?))
                        args.Add(String.IsNullOrEmpty(val) ? null : (object)Convert.ChangeType(val, typeof(int)));
                    else if (p.ParameterType == typeof(long?))
                        args.Add(String.IsNullOrEmpty(val) ? null : (object)Convert.ChangeType(val, typeof(long)));
                    else if (p.ParameterType == typeof(float?))
                        args.Add(String.IsNullOrEmpty(val) ? null : (object)Convert.ChangeType(val, typeof(float)));
                    else if (p.ParameterType == typeof(double?))
                        args.Add(String.IsNullOrEmpty(val) ? null : (object)Convert.ChangeType(val, typeof(double)));
                    else if (p.ParameterType == typeof(DateTime?))
                        args.Add(String.IsNullOrEmpty(val) ? null : (object)Convert.ChangeType(val, typeof(DateTime)));
                    else
                        args.Add(String.IsNullOrEmpty(val) ? null : (object)Convert.ChangeType(val, p.ParameterType));
                }
                dataObjects.Add((T)ctor.Invoke(args.ToArray()));
            }
            return dataObjects.AsEnumerable();
        }

        public static List<List<string>> Tabulate<T>(this IEnumerable<T> dataObjects)
        {
            if (dataObjects.Count() == 0) return null;
            var ctor = dataObjects.ElementAt(0).GetType().GetConstructors().Single();
            var pars = ctor.GetParameters();
            List<List<string>> dataTable = new List<List<string>>();
            List<string> header = new List<string>();
            foreach (var p in pars)
                header.Add(p.Name);
            dataTable.Add(header);
            foreach (var entity in dataObjects)
            {
                List<string> record = new List<string>();
                foreach (var p in pars)
                {
                    object val = entity.GetType().GetProperty(p.Name).GetValue(entity, null);
                    record.Add(val != null ? val.ToString() : String.Empty);
                }
                dataTable.Add(record);
            }
            return dataTable;
        }

        public static List<List<string>> Rename(this List<List<string>> dataTable, params string[] names)
        {
            if (dataTable[0].Count != names.Count()) throw new Exception("Number of columns dismatch.");
            dataTable[0] = names.ToList();
            return dataTable;
        }
    }
}
