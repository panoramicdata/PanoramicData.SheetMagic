# PanoramicData.SheetMagic
Easily save/load from/to Excel (XLSX) documents using generics in C#

Example:

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
using (var workbook = new MagicSpreadsheet(fileInfo))
{
	workbook.AddSheet(things);
	workbook.Save();
}