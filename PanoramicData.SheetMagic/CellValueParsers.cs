namespace PanoramicData.SheetMagic;

/// <summary>
/// Parsers for untyped cell value reading
/// </summary>
internal static class CellValueParsers
{
	internal static string GetSharedStringValue(SharedStringTablePart? stringTable, string? cellValueText)
	{
		if (stringTable?.SharedStringTable is null || cellValueText == null)
		{
			return string.Empty;
		}

		return stringTable.SharedStringTable.ElementAt(int.Parse(cellValueText)).InnerText;
	}

	internal static bool? ParseBooleanValue(string? cellValueText)
		=> cellValueText switch
		{
			"1" or "true" => true,
			"0" or "false" => false,
			_ => null,
		};

	internal static double? ParseNumberValue(string? cellValueText)
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

	internal static object? GetStringOrInfinityValue(string? cellValueText, Cell cell)
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
