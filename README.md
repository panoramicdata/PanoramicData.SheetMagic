# PanoramicData.SheetMagic

[![Codacy Badge](https://app.codacy.com/project/badge/Grade/70d9929b4d3c4d8ab2d69c5209a29b6e)](https://www.codacy.com/gh/panoramicdata/PanoramicData.SheetMagic/dashboard?utm_source=github.com&amp;utm_medium=referral&amp;utm_content=panoramicdata/PanoramicData.SheetMagic&amp;utm_campaign=Badge_Grade)
![Commit Activity](https://img.shields.io/github/commit-activity/m/panoramicdata/PanoramicData.SheetMagic)

Easily save/load from/to Excel (XLSX) documents using generics in C#

## Writing to a file

```c#
// Write a list of items to an XLSX file
var things = new List<Thing>
{
	new Thing
	{
		PropertyA = "Value 1",
		PropertyB = 1
	},
	new Thing
	{
		PropertyA = "Value 2",
		PropertyB = 2
	},
};
var fileInfo = new FileInfo($"Output {DateTime.UtcNow:yyyyMMddTHHmmss}Z.xlsx");
using var workbook = new MagicSpreadsheet(fileInfo);
workbook.AddSheet(things);
workbook.Save();
```

## Reading from a file

```c#
// Read a list of items from an XLSX file
using var workbook = new MagicSpreadsheet(fileInfo);
workbook.Load();
// Use default worksheet
var cars = workbook.GetList<Car>();
// Use a different worksheet
var animals = workbook.GetList<Animal>("Animals");
