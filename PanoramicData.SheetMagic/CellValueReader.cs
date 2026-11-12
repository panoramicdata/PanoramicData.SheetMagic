namespace PanoramicData.SheetMagic;

/// <summary>
/// Untyped cell value reading
/// </summary>
internal static class CellValueReader
{
	internal static string GetCellValueString(Cell cell, SharedStringTablePart? stringTable)
	{
		var cellValueText = cell.CellValue?.Text;

		if (cell.DataType != null && (CellValues)cell.DataType == CellValues.SharedString)
		{
			return GetSharedStringValue(stringTable, cellValueText);
		}

		return cellValueText ?? cell.InnerText;
	}


	internal static object? GetCellValueByDataType(Cell cell, SharedStringTablePart? stringTable, string? cellValueText)
	{
		// DocumentFormat.OpenXml 3.x turned CellValues from an enum into a struct, so its
		// members are no longer compile-time constants and cannot be used as switch patterns.
		var dataType = (CellValues)cell.DataType!;

		if (dataType == CellValues.SharedString)
		{
			return GetSharedStringValue(stringTable, cellValueText);
		}

		if (dataType == CellValues.Boolean)
		{
			return ParseBooleanValue(cellValueText);
		}

		if (dataType == CellValues.Number)
		{
			return ParseNumberValue(cellValueText);
		}

		if (dataType == CellValues.Date)
		{
			return DateTime.Parse(cellValueText!);
		}

		if (dataType == CellValues.Error || dataType == CellValues.String || dataType == CellValues.InlineString)
		{
			return GetStringOrInfinityValue(cellValueText, cell);
		}

		throw new NotSupportedException($"Unsupported data type {cell.DataType?.Value.ToString() ?? "None"}");
	}
}
