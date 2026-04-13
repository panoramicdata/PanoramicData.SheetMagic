[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

[![NuGet version](https://img.shields.io/nuget/v/PanoramicData.SheetMagic.svg)](https://www.nuget.org/packages/PanoramicData.SheetMagic/)

# PanoramicData.SheetMagic

[![Codacy Badge](https://app.codacy.com/project/badge/Grade/70d9929b4d3c4d8ab2d69c5209a29b6e)](https://www.codacy.com/gh/panoramicdata/PanoramicData.SheetMagic/dashboard?utm_source=github.com&amp;utm_medium=referral&amp;utm_content=panoramicdata/PanoramicData.SheetMagic&amp;utm_campaign=Badge_Grade)
![Commit Activity](https://img.shields.io/github/commit-activity/m/panoramicdata/PanoramicData.SheetMagic)
![.NET Version](https://img.shields.io/badge/.NET-9.0-512BD4)

Easily save/load data to/from Excel (XLSX) documents using strongly-typed C# classes.

## Requirements

- **.NET 9.0** - This library targets .NET 9.0 only

## Installation

```bash
dotnet add package PanoramicData.SheetMagic
```

## Features

- ? **Strongly-typed** - Work with your own C# classes
- ? **Simple API** - Easy to read and write XLSX files
- ? **Multiple sheets** - Add and read multiple worksheets
- ? **Styling support** - Apply table styles to your data
- ? **Extended properties** - Support for dynamic properties via `Extended<T>`
- ? **Streams and files** - Work with both `FileInfo` and `Stream` objects
- ? **Type safe** - Full support for common .NET types including nullable types

## Quick Start

### Writing to a file

```csharp
using PanoramicData.SheetMagic;

// Define your class
public class Thing
{
    public string PropertyA { get; set; }
    public int PropertyB { get; set; }
}

// Create some data
var things = new List<Thing>
{
    new Thing { PropertyA = "Value 1", PropertyB = 1 },
    new Thing { PropertyA = "Value 2", PropertyB = 2 },
};

// Write to Excel file
var fileInfo = new FileInfo($"Output {DateTime.UtcNow:yyyyMMddTHHmmss}Z.xlsx");
using var workbook = new MagicSpreadsheet(fileInfo);
workbook.AddSheet(things);
workbook.Save();
```

### Reading from a file

```csharp
using PanoramicData.SheetMagic;

// Read from Excel file
using var workbook = new MagicSpreadsheet(fileInfo);
workbook.Load();

// Read from default worksheet (first sheet)
var cars = workbook.GetList<Car>();

// Read from a specific worksheet by name
var animals = workbook.GetList<Animal>("Animals");
```

## Advanced Features

### Working with Streams

```csharp
// Write to a stream
using var stream = new MemoryStream();
using (var workbook = new MagicSpreadsheet(stream))
{
    workbook.AddSheet(data);
    workbook.Save();
}

// Read from a stream
stream.Position = 0;
using var workbook = new MagicSpreadsheet(stream);
workbook.Load();
var items = workbook.GetList<MyClass>();
```

### Multiple Sheets

```csharp
using var workbook = new MagicSpreadsheet(fileInfo);
workbook.AddSheet(cars, "Cars");
workbook.AddSheet(animals, "Animals");
workbook.AddSheet(products, "Products");
workbook.Save();
```

### Applying Table Styles

```csharp
var options = new AddSheetOptions
{
    TableOptions = new TableOptions
    {
        Name = "MyTable",
        DisplayName = "MyTable1",
        XlsxTableStyle = XlsxTableStyle.TableStyleMedium2,
      ShowRowStripes = true,
        ShowColumnStripes = false,
        ShowFirstColumn = false,
   ShowLastColumn = false
    }
};

workbook.AddSheet(data, "StyledSheet", options);
```

### Conditional Formatting

Conditional formatting is configured through `AddSheetOptions.ConditionalFormats`.
The object model is intentionally close to Excel's own configuration model:

- `ConditionalFormat` selects one or more output columns, or all columns when `ColumnNames` is omitted.
- `ConditionalFormatRule` describes one Excel rule such as `CellIs`, `ContainsBlanks`, or `ContainsErrors`.
- `ConditionalFormatStyle` defines the differential format Excel applies when the rule matches.

`ColumnNames` must match the final header text written to Excel.
If `PropertyHeaders` is set, use those values.
Otherwise use the `Description` attribute value or the property name.

```csharp
using System.Drawing;
using PanoramicData.SheetMagic;

var rows = new[]
{
    new ReportRow { Name = null, Description = "Missing name", Score = null },
    new ReportRow { Name = "Bravo", Description = "High score", Score = 9 },
    new ReportRow { Name = "Charlie", Description = null, Score = 4 }
}.ToList();

var options = new AddSheetOptions
{
    ConditionalFormats =
    [
        new ConditionalFormat
        {
            ColumnNames = ["Name", "Description"],
            Rules =
            [
                new ConditionalFormatRule
                {
                    RuleType = ConditionalFormatRuleType.ContainsBlanks,
                    Style = new ConditionalFormatStyle
                    {
                        BackgroundColor = Color.Red
                    }
                }
            ]
        },
        new ConditionalFormat
        {
            ColumnNames = ["Score"],
            Rules =
            [
                new ConditionalFormatRule
                {
                    RuleType = ConditionalFormatRuleType.ContainsBlanks,
                    Style = new ConditionalFormatStyle
                    {
                        BackgroundColor = Color.Red
                    }
                },
                new ConditionalFormatRule
                {
                    RuleType = ConditionalFormatRuleType.CellIs,
                    Operator = ConditionalFormatOperator.GreaterThan,
                    Formula = "5",
                    Style = new ConditionalFormatStyle
                    {
                        FontColor = Color.Green,
                        FontWeight = FontWeight.Bold
                    }
                }
            ]
        },
        new ConditionalFormat
        {
            Rules =
            [
                new ConditionalFormatRule
                {
                    RuleType = ConditionalFormatRuleType.ContainsErrors,
                    Style = new ConditionalFormatStyle
                    {
                        FontWeight = FontWeight.Bold
                    }
                }
            ]
        }
    ]
};

using var workbook = new MagicSpreadsheet(new FileInfo("ConditionalFormatting.xlsx"));
workbook.AddSheet(rows, "Report", options);
workbook.Save();

public sealed class ReportRow
{
    public string? Name { get; set; }

    public string? Description { get; set; }

    public int? Score { get; set; }
}
```

Supported rule types currently include:

- `CellIs`
- `Expression`
- `ContainsBlanks`
- `NotContainsBlanks`
- `ContainsErrors`
- `NotContainsErrors`
- `ContainsText`
- `NotContainsText`
- `BeginsWith`
- `EndsWith`
- `DuplicateValues`
- `UniqueValues`
- `Top10`
- `AboveAverage`

### Custom Property Headers

Use the `Description` attribute to customize column headers:

```csharp
using System.ComponentModel;

public class Employee
{
    public int Id { get; set; }
    
    [Description("Full Name")]
    public string Name { get; set; }
    
    [Description("Hire Date")]
    public DateTime HireDate { get; set; }
}
```

### Property Filtering

```csharp
// Include only specific properties
var options = new AddSheetOptions
{
 IncludeProperties = new[] { "Name", "Age", "City" }
};
workbook.AddSheet(people, "Filtered", options);

// Exclude specific properties
var options = new AddSheetOptions
{
    ExcludeProperties = new[] { "InternalId", "Password" }
};
workbook.AddSheet(users, "Public", options);
```

### Extended Properties (Dynamic Properties)

Work with objects that have both strongly-typed and dynamic properties:

```csharp
var extendedData = new List<Extended<MyClass>>
{
    new Extended<MyClass>(
        new MyClass { Id = 1, Name = "Item 1" },
    new Dictionary<string, object?>
        {
    { "DynamicProp1", "Value1" },
        { "DynamicProp2", 42 }
 }
    )
};

workbook.AddSheet(extendedData);
workbook.Save();

// Reading extended properties
var loadedData = workbook.GetExtendedList<MyClass>();
foreach (var item in loadedData)
{
    Console.WriteLine($"{item.Item.Name}");
    foreach (var prop in item.Properties)
    {
        Console.WriteLine($"  {prop.Key}: {prop.Value}");
    }
}
```

### Supported Types

- Primitives: `int`, `long`, `short`, `uint`, `ulong`, `ushort`
- Floating point: `float`, `double`, `decimal`
- Boolean: `bool`
- Dates: `DateTime`, `DateTimeOffset`
- Strings: `string`
- Enums (stored as text)
- Lists: `List<string>` (with configurable delimiter)
- All nullable versions of the above

### Options

Configure behavior with the `Options` class:

```csharp
var options = new Options
{
    StopProcessingOnFirstEmptyRow = true,
IgnoreUnmappedProperties = true,
    EmptyRowInterpretedAsNull = false,
  LoadNullExtendedProperties = true,
    ListSeparator = ";"
};

using var workbook = new MagicSpreadsheet(fileInfo, options);
```

## Known Limitations

- **JObject Support**: Direct `JObject` serialization is not yet supported. Use `Extended<object>` instead.
- **Nested Complex Objects**: Properties of type `List<ComplexType>` cannot be loaded from Excel (though they can be saved as delimited strings).
- **Large Integer Precision**: Excel stores all numbers as doubles, so very large `Int64`/`UInt64` values (near `MaxValue`) may lose precision.
- **Special Values**: `double.NaN` and `null` nullable types are stored as empty strings in Excel.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

See the [LICENSE](LICENSE) file for details.
