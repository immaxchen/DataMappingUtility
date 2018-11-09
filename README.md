# DataMappingUtility
C# utility for conversion between tabular data and anonymous typed / compatible class objects

Target .NET Framework 3.5

Need to reference EPPlus for Excel I/O

# Example

create anonymous typed objects on the fly

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

or using a compatible class definition

```c#
var users = table.Generate(new User());
```
