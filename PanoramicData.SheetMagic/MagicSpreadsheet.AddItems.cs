using PanoramicData.SheetMagic.Extensions;

namespace PanoramicData.SheetMagic;

/// <summary>
/// Methods for processing and adding items to sheets
/// </summary>
public partial class MagicSpreadsheet
{
	private SheetLayout AddItems<T>(
		List<T> items,
		AddSheetOptions addSheetOptions,
		Type type,
		bool isExtended,
		SheetData sheetData)
	{
		var (basicType, keyHashSet) = DetermineTypeAndCollectKeys(items, addSheetOptions, type, isExtended);

		var propertyList = GetFilteredAndOrderedProperties(basicType, addSheetOptions);
		var keyList = SortKeyList(keyHashSet, addSheetOptions);

		var totalColumnCount = (uint)(propertyList.Count + keyList.Count);
		var columnConfigurations = CreateColumns(totalColumnCount);

		AddHeaderRow(sheetData, propertyList, keyList, addSheetOptions);
		AddDataRows(items, addSheetOptions, type, isExtended, sheetData, propertyList, keyList);

		return new SheetLayout(propertyList, columnConfigurations, keyList, totalColumnCount);
	}

	private static Columns CreateColumns(uint totalColumnCount)
	{
		var columnConfigurations = new Columns();
		for (var n = 0; n < totalColumnCount; n++)
		{
			_ = columnConfigurations.AppendChild(new Column { BestFit = true });
		}
		return columnConfigurations;
	}

	private static void AddHeaderRow(
		SheetData sheetData,
		List<PropertyInfo> propertyList,
		List<string> keyList,
		AddSheetOptions addSheetOptions)
	{
		uint rowIndex = 1;
		var row = new Row { RowIndex = rowIndex };
		_ = sheetData.AppendChild(row);
		var writer = new RowWriter(row, rowIndex);

		var headers = GetHeaders(propertyList, addSheetOptions);
		AddHeaderCells(writer, headers);
		AddHeaderCells(writer, keyList);
	}

	private static string[] GetHeaders(List<PropertyInfo> propertyList, AddSheetOptions addSheetOptions)
		=> addSheetOptions.PropertyHeaders?.Length == propertyList.Count
			? addSheetOptions.PropertyHeaders
			: [.. propertyList.Select(p => p.GetPropertyDescription() ?? p.Name)];

	private static void AddHeaderCells(RowWriter writer, IEnumerable<string> headers)
	{
		foreach (var header in headers)
		{
			writer.AppendValue(header ?? string.Empty);
		}
	}
}
