namespace PanoramicData.SheetMagic;

/// <summary>
/// Untyped cell value reading
/// </summary>
public partial class MagicSpreadsheet
{
	private static string GetCellValueString(Cell cell, SharedStringTablePart? stringTable)
	{
		var cellValueText = cell.CellValue?.Text;

		if (cell.DataType != null && (CellValues)cell.DataType == CellValues.SharedString)
		{
			return GetSharedStringValue(stringTable, cellValueText);
		}

		return cellValueText ?? cell.InnerText;
	}

	private object? GetCellValueDirect(Cell cell, SharedStringTablePart? stringTable)
	{
		var cellValueText = cell.CellValue?.Text;
		if (cell.DataType == null)
		{
			// Check whether there is a built-in style set
			return GetCellFormatFromStyle(cell) is string text
				? text
				: (object?)cellValueText;
		}

		return GetCellValueByDataType(cell, stringTable, cellValueText);
	}

	private static object? GetCellValueByDataType(Cell cell, SharedStringTablePart? stringTable, string? cellValueText)
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

	private static string GetSharedStringValue(SharedStringTablePart? stringTable, string? cellValueText)
	{
		if (stringTable?.SharedStringTable is null || cellValueText == null)
		{
			return string.Empty;
		}

		return stringTable.SharedStringTable.ElementAt(int.Parse(cellValueText)).InnerText;
	}

	private static bool? ParseBooleanValue(string? cellValueText)
		=> cellValueText switch
		{
			"1" or "true" => true,
			"0" or "false" => false,
			_ => null,
		};

	private static double? ParseNumberValue(string? cellValueText)
	{
		if (cellValueText == null)
		{
			return null;
		}

		// Handle special double values that cannot be parsed normally
		return cellValueText switch
		{
			"Infinity" => double.PositiveInfinity,
			"-Infinity" => double.NegativeInfinity,
			"NaN" => double.NaN,
			_ => double.Parse(cellValueText),
		};
	}

	private static object? GetStringOrInfinityValue(string? cellValueText, Cell cell)
	{
		// For InlineString cells, get the actual text value
		var textValue = cellValueText;
		if (textValue == null && cell.DataType != null && cell.DataType == CellValues.InlineString)
		{
			// Extract text from InlineString element
			var inlineString = cell.Elements<InlineString>().FirstOrDefault();
			textValue = inlineString?.Text?.Text;
		}

		textValue ??= cell.InnerText;

		// Handle special values for string cells as well (since Infinity values are stored as text)
		return textValue switch
		{
			"Infinity" => double.PositiveInfinity,
			"-Infinity" => double.NegativeInfinity,
			_ => textValue,
		};
	}
}
