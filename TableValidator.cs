using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DataMappingUtility
{
    class TableValidator
    {
        private List<List<string>> Table { get; set; }
        private Dictionary<string, TableField> Fields { get; set; }
        private Dictionary<string[], TableCompositeField> CompositeFields { get; set; }

        public TableValidator(List<List<string>> dataTable)
        {
            this.Table = dataTable;
            this.Fields = new Dictionary<string, TableField>();
            this.CompositeFields = new Dictionary<string[], TableCompositeField>(new StringArrayEqualityComparer());
        }

        public TableField Field(string colName)
        {
            if (!Fields.ContainsKey(colName))
            {
                var index = Table[0].IndexOf(colName);
                if (index == -1) throw new KeyNotFoundException("Incorrect Column Name.");
                Fields.Add(colName, new TableField(this, index, colName));
            }
            return Fields[colName];
        }

        public TableCompositeField CompositeField(params string[] colNames)
        {
            if (!CompositeFields.ContainsKey(colNames))
            {
                var indices = colNames.Select(o => Table[0].IndexOf(o)).ToArray();
                if (indices.Contains(-1)) throw new KeyNotFoundException("Incorrect Column Names.");
                var grpName = "(" + string.Join(", ", colNames) + ")";
                CompositeFields.Add(colNames, new TableCompositeField(this, indices, grpName));
            }
            return CompositeFields[colNames];
        }

        public string Validate()
        {
            var sb = new StringBuilder();
            for (int i = 1; i < Table.Count; i++)
            {
                var row = Table[i];

                foreach (var col in Fields.Values)
                {
                    foreach (var check in col.Checkers)
                    {
                        var message = check(row);
                        if (message != null) sb.AppendLine(string.Format("[ Row#{0:D6} ] {1}", i + 1, message));
                    }
                }

                foreach (var grp in CompositeFields.Values)
                {
                    foreach (var check in grp.Checkers)
                    {
                        var message = check(row);
                        if (message != null) sb.AppendLine(string.Format("[ Row#{0:D6} ] {1}", i + 1, message));
                    }
                }
            }
            return sb.ToString();
        }


        public class TableField
        {
            public TableValidator Parent { get; set; }
            public int Index { get; set; }
            public string Name { get; set; }
            public List<Func<List<string>, string>> Checkers { get; set; }

            public TableField(TableValidator parent, int index, string name)
            {
                this.Parent = parent;
                this.Index = index;
                this.Name = name;
                this.Checkers = new List<Func<List<string>, string>>();
            }

            public void AddConstraint(Func<string, string> constraint)
            {
                Checkers.Add(o => constraint(o[Index]));
            }

            public void AddComparator(string targetColName, Func<string, string, string> comparator)
            {
                Checkers.Add(o => comparator(o[Index], o[Parent.Fields[targetColName].Index]));
            }

            public static bool IsNullOrSpace(string s)
            {
                return string.IsNullOrEmpty((s ?? string.Empty).Trim());
            }

            public static long? ParseInteger(string s)
            {
                long result;
                return long.TryParse(s, out result) ? result : (long?)null;
            }

            public static double? ParseDouble(string s)
            {
                double result;
                return double.TryParse(s, out result) ? result : (double?)null;
            }

            public void IsRequired()
            {
                AddConstraint(s =>
                {
                    return IsNullOrSpace(s) ? Name + " cannot be empty" : null;
                });
            }

            public void IsUnique()
            {
                var pool = new HashSet<string>();

                AddConstraint(s =>
                {
                    return pool.Add(s) ? null : string.Format("{0} should be unique, found duplicate: {1}", Name, s);
                });
            }

            public void IsInteger()
            {
                AddConstraint(s =>
                {
                    if (IsNullOrSpace(s)) return null;
                    return ParseInteger(s) == null ? string.Format("{0} should be an integer, got: {1}", Name, s) : null;
                });
            }

            public void IsNumeric()
            {
                AddConstraint(s =>
                {
                    if (IsNullOrSpace(s)) return null;
                    return ParseDouble(s) == null ? string.Format("{0} should be a number, got: {1}", Name, s) : null;
                });
            }

            public void IsGreaterThan(double value)
            {
                AddConstraint(s =>
                {
                    var result = ParseDouble(s);
                    return result == null || result > value ? null : string.Format("{0} should be greater than {1}, got: {2}", Name, value, s);
                });
            }

            public void IsLessThan(double value)
            {
                AddConstraint(s =>
                {
                    var result = ParseDouble(s);
                    return result == null || result < value ? null : string.Format("{0} should be less than {1}, got: {2}", Name, value, s);
                });
            }

            public void IsIn(params string[] values)
            {
                AddConstraint(s =>
                {
                    return IsNullOrSpace(s) || values.Contains(s) ? null : string.Format("{0} should be one of: {{{1}}}, got: {2}", Name, string.Join(", ", values), s);
                });
            }

            public void IsGreaterThan(string targetColName)
            {
                AddComparator(targetColName, (s1, s2) =>
                {
                    var v1 = ParseDouble(s1);
                    var v2 = ParseDouble(s2);
                    return v1 == null || v2 == null || v1 > v2 ? null : string.Format("{0} should be greater than {1}, got: {2}, {3}", Name, targetColName, s1, s2);
                });
            }
        }


        public class TableCompositeField
        {
            public TableValidator Parent { get; set; }
            public int[] Indices { get; set; }
            public string Name { get; set; }
            public List<Func<List<string>, string>> Checkers { get; set; }

            public TableCompositeField(TableValidator parent, int[] indices, string name)
            {
                this.Parent = parent;
                this.Indices = indices;
                this.Name = name;
                this.Checkers = new List<Func<List<string>, string>>();
            }

            public void AddConstraint(Func<string[], string> constraint)
            {
                Checkers.Add(o => constraint(o.Where((v, i) => Indices.Contains(i)).ToArray()));
            }

            public void IsRequired()
            {
                AddConstraint(a =>
                {
                    return a.All(s => TableField.IsNullOrSpace(s)) ? Name + " cannot all be empty" : null;
                });
            }

            public void IsUnique()
            {
                var pool = new HashSet<string[]>(new StringArrayEqualityComparer());

                AddConstraint(a =>
                {
                    return pool.Add(a) ? null : string.Format("{0} should be unique, found duplicate: ({1})", Name, string.Join(", ", a));
                });
            }
        }


        public class StringArrayEqualityComparer : IEqualityComparer<string[]>
        {
            public bool Equals(string[] x, string[] y)
            {
                return x.SequenceEqual(y);
            }

            public int GetHashCode(string[] obj)
            {
                unchecked
                {
                    int hash = 17;
                    foreach (var o in obj) hash = hash * 23 + (o ?? string.Empty).GetHashCode();
                    return hash;
                }
            }
        }
    }
}
