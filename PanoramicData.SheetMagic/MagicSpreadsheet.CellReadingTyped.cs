namespace PanoramicData.SheetMagic;

/// <summary>
/// Typed cell value reading
/// </summary>
public partial class MagicSpreadsheet
{
	private static object? GetCellValue<T>(Cell cell, SharedStringTablePart? stringTable)
	{
		var cellValueText = cell.CellValue?.Text;

		if (cell.DataType == null)
		{
			return GetCellValueWithoutDataType<T>(cell, cellValueText);
		}

		return GetCellValueWithDataType<T>(cell, stringTable, cellValueText);
	}

	private static object? GetCellValueWithoutDataType<T>(Cell cell, string? cellValueText)
	{
		return typeof(T).Name switch
		{
			"Int32" => ParseOrThrow<int>(cellValueText, int.TryParse, cell, "integer"),
			"Int64" => ParseOrThrow<long>(cellValueText, long.TryParse, cell, "integer"),
			"Double" => ParseDouble(cellValueText, cell),
			"Single" => ParseFloat(cellValueText, cell),
			"Boolean" => ParseOrThrow<bool>(cellValueText, bool.TryParse, cell, "bool"),
			"Nullable`1<Boolean>" => ParseNullableBool(cellValueText),
			"String" or "Object" => cellValueText,
			_ => throw new NotSupportedException($"Unsupported data type {typeof(T).Name}"),
		};
	}

	private static object? ParseOrThrow<TValue>(string? input, TryParseDelegate<TValue> tryParse, Cell cell, string typeName)
	{
		if (tryParse(input, out var value))
		{
			return value;
		}
		throw new FormatException($"Could not convert cell {cell.CellReference} to {typeName}.");
	}

	private delegate bool TryParseDelegate<TValue>(string? input, out TValue value);

	private static double? ParseDouble(string? cellValueText, Cell cell)
	{
		// Handle special Infinity values
		var infinity = TryParseInfinity(cellValueText);
		if (infinity is not null)
		{
			return infinity;
		}

		if (int.TryParse(cellValueText, out var doubleValue))
		{
			return doubleValue;
		}

		throw new FormatException($"Could not convert cell {cell.CellReference} to a double.");
	}

	private static float? ParseFloat(string? cellValueText, Cell cell)
	{
		// Handle special Infinity values
		var infinity = TryParseInfinity(cellValueText);
		if (infinity is not null)
		{
			return (float)infinity.Value;
		}

		if (float.TryParse(cellValueText, out var floatValue))
		{
			return floatValue;
		}

		throw new FormatException($"Could not convert cell {cell.CellReference} to a float.");
	}

	/// <summary>
	/// Recognises the text that <see cref="CreateNumericOrSpecialCell"/> writes for infinite values,
	/// which Excel itself has no numeric representation for.
	/// </summary>
	/// <returns>The infinity the text denotes, or null if it denotes no infinity.</returns>
	private static double? TryParseInfinity(string? text)
		=> text switch
		{
			"Infinity" => double.PositiveInfinity,
			"-Infinity" => double.NegativeInfinity,
			_ => null
		};

	private static bool? ParseNullableBool(string? cellValueText)
	{
		if (cellValueText == "NULL" || cellValueText == string.Empty)
		{
			return null;
		}

		if (bool.TryParse(cellValueText, out var boolValue))
		{
			return boolValue;
		}

		throw new FormatException($"Could not convert cell value to a bool.");
	}

	private static object? GetCellValueWithDataType<T>(Cell cell, SharedStringTablePart? stringTable, string? cellValueText)
	{
		// See GetCellValueByDataType: CellValues is a struct in OpenXml 3.x, not an enum.
		if (cell.DataType is null)
		{
			return null;
		}

		var dataType = (CellValues)cell.DataType;

		if (dataType == CellValues.SharedString)
		{
			return GetSharedStringValueTyped<T>(stringTable, cellValueText);
		}

		if (dataType == CellValues.Boolean)
		{
			return ParseBooleanValueTyped(cellValueText);
		}

		if (dataType == CellValues.Number)
		{
			return ParseNumberTyped(cellValueText);
		}

		return GetTextualCellValueWithDataType<T>(cell, dataType, cellValueText);
	}

	private static object? GetTextualCellValueWithDataType<T>(Cell cell, CellValues dataType, string? cellValueText)
	{
		if (dataType == CellValues.Date)
		{
			return ParseDateTyped(cellValueText);
		}

		if (dataType == CellValues.Error || dataType == CellValues.String || dataType == CellValues.InlineString)
		{
			return ParseStringOrInfinityTyped<T>(cellValueText, cell);
		}

		throw new NotSupportedException($"Unsupported data type {cell.DataType?.Value.ToString() ?? "None"}");
	}

	private static object? GetSharedStringValueTyped<T>(SharedStringTablePart? stringTable, string? cellValueText)
	{
		if (stringTable?.SharedStringTable is null || cellValueText == null)
		{
			throw new FormatException("SharedStringTable or cell value text is null for SharedString type");
		}

		var stringTableIndex = int.Parse(cellValueText);
		var sharedStringElement = stringTable.SharedStringTable.ElementAt(stringTableIndex);
		// Handle special Infinity values for object-typed cells
		return TryParseInfinityForObject<T>(sharedStringElement.InnerText) ?? (object)sharedStringElement.InnerText;
	}

	/// <summary>
	/// Infinities are written as text, so they only read back as numbers where the target type is
	/// permissive enough to take one.
	/// </summary>
	private static double? TryParseInfinityForObject<T>(string? text)
		=> typeof(T).Name == "Object"
			? TryParseInfinity(text)
			: null;

	private static bool? ParseBooleanValueTyped(string? cellValueText)
		=> ParseBooleanValue(cellValueText);

	private static double ParseNumberTyped(string? cellValueText)
	{
		if (cellValueText == null)
		{
			throw new FormatException("Cell value text is null for Number type");
		}

		return double.Parse(cellValueText);
	}

	private static DateTime ParseDateTyped(string? cellValueText)
	{
		if (cellValueText == null)
		{
			throw new FormatException("Cell value text is null for Date type");
		}

		return DateTime.Parse(cellValueText);
	}

	private static object? ParseStringOrInfinityTyped<T>(string? cellValueText, Cell cell)
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

	private static string GetCellText(string? cellValueText, Cell cell)
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
