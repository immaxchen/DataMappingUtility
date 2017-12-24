# DataMappingUtility
C# utility for converting between tabular data and anonymous typed objects

Target .NET Framework 3.5

Need to reference EPPlus for Excel I/O

# Example

```c#
using DataMappingUtility;

var table = DataIO.ReadExcel("users.xlsx");
var users = table.Generate(new
    {
        UserId = default(int?),
        Name = default(string),
        Gender = default(string),
        Birthday = default(DateTime?),
    });

var boys = users.Where(x => x.Gender == "Male");
```
