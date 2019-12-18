using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
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
                    {
                        if (sheet.Cells[i, j].Value == null)
                            record.Add(String.Empty);
                        else
                            record.Add(sheet.Cells[i, j].Value.ToString());
                    }
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

        public static List<List<string>> ReadDataTable(DataTable table)
        {
            return new List<List<string>>() {
                table.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToList()
            }.Concat(
                table.AsEnumerable().Select(x => x.ItemArray.Select(o => o.ToString()).ToList()).ToList()
            ).ToList();
        }

        public static DataTable ToDataTable(this List<List<string>> dataTable)
        {
            var table = new DataTable();
            if (dataTable.Count == 0) return table;
            foreach (string column in dataTable[0])
                table.Columns.Add(column);
            if (dataTable.Count == 1) return table;
            foreach (var record in dataTable.Skip(1))
                table.LoadDataRow((object[])record.ToArray(), true);
            return table;
        }

        public static List<List<string>> ReadMdb(string filepath, string tablename)
        {
            var table = new DataTable();
            using (var connection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filepath))
            using (var adapter = new OleDbDataAdapter("SELECT * FROM " + tablename, connection))
                adapter.Fill(table);
            return ReadDataTable(table);
        }
    }

    public static class DataOOP
    {
        public static IEnumerable<T> Generate<T>(this List<List<string>> dataTable, T dataModel, int ctorIndex = 0)
        {
            if (dataTable.Count == 0) return null;
            var ctor = dataModel.GetType().GetConstructors()[ctorIndex];
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

        public static List<List<string>> Tabulate<T>(this IEnumerable<T> dataObjects, int ctorIndex = 0)
        {
            if (dataObjects.Count() == 0) return null;
            var ctor = dataObjects.ElementAt(0).GetType().GetConstructors()[ctorIndex];
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
            if (dataTable[0].Count != names.Count()) throw new Exception("Number of column names mismatch.");
            dataTable[0] = names.ToList();
            return dataTable;
        }

        public static List<List<string>> Rename(this List<List<string>> dataTable, Dictionary<string, string> mapper)
        {
            foreach (var entry in mapper)
                if (dataTable[0].Contains(entry.Key)) dataTable[0][dataTable[0].IndexOf(entry.Key)] = entry.Value;
                else throw new Exception("The given column name was not present in the table.");
            return dataTable;
        }
    }
}
