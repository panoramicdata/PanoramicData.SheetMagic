namespace PanoramicData.SheetMagic;

/// <summary>
/// Parsers used by typed cell value reading
/// </summary>
internal static class TypedCellParsers
{
	internal static bool? ParseBooleanValueTyped(string? cellValueText)
		=> ParseBooleanValue(cellValueText);

	internal static double ParseNumberTyped(string? cellValueText)
	{
		if (cellValueText == null)
		{
			throw new FormatException("Cell value text is null for Number type");
		}

		return double.Parse(cellValueText);
	}

	internal static DateTime ParseDateTyped(string? cellValueText)
	{
		if (cellValueText == null)
		{
			throw new FormatException("Cell value text is null for Date type");
		}

		return DateTime.Parse(cellValueText);
	}

	internal static object? ParseStringOrInfinityTyped<T>(string? cellValueText, Cell cell)
	{
		var textValue = GetCellText(cellValueText, cell);

		// Handle special Infinity values for object type
		var infinity = TryParseInfinityForObject<T>(textValue);
		if (infinity is not null)
		{
			return infinity;
		}

		try
		{
			return (T)Convert.ChangeType(textValue, typeof(T));
		}
		catch (FormatException)
		{
			return null;
		}
	}

	internal static string GetCellText(string? cellValueText, Cell cell)
	{
		if (cellValueText is not null)
		{
			return cellValueText;
		}

		// For InlineString cells, get the actual text value
		if (cell.DataType is not null && cell.DataType == CellValues.InlineString)
		{
			// Extract text from InlineString element
			var inlineText = cell.Elements<InlineString>().FirstOrDefault()?.Text?.Text;
			if (inlineText is not null)
			{
				return inlineText;
			}
		}

		return cell.InnerText;
	}
}
