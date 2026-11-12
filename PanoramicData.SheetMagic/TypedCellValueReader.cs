namespace PanoramicData.SheetMagic;

/// <summary>
/// Typed cell value reading for cells without a data type
/// </summary>
internal static class TypedCellValueReader
{
	internal static object? GetCellValue<T>(Cell cell, SharedStringTablePart? stringTable)
	{
		var cellValueText = cell.CellValue?.Text;

		if (cell.DataType == null)
		{
			return GetCellValueWithoutDataType<T>(cell, cellValueText);
		}

		return GetCellValueWithDataType<T>(cell, stringTable, cellValueText);
	}

	internal static object? GetCellValueWithoutDataType<T>(Cell cell, string? cellValueText)
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

	internal static object? ParseOrThrow<TValue>(string? input, TryParseDelegate<TValue> tryParse, Cell cell, string typeName)
	{
		if (tryParse(input, out var value))
		{
			return value;
		}
		throw new FormatException($"Could not convert cell {cell.CellReference} to {typeName}.");
	}

	internal delegate bool TryParseDelegate<TValue>(string? input, out TValue value);

	internal static double? ParseDouble(string? cellValueText, Cell cell)
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

	internal static float? ParseFloat(string? cellValueText, Cell cell)
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
	/// Recognises the text that <see cref="CellOperationHelpers.CreateNumericOrSpecialCell"/> writes for infinite values,
	/// which Excel itself has no numeric representation for.
	/// </summary>
	/// <returns>The infinity the text denotes, or null if it denotes no infinity.</returns>
	internal static double? TryParseInfinity(string? text)
		=> text switch
		{
			"Infinity" => double.PositiveInfinity,
			"-Infinity" => double.NegativeInfinity,
			_ => null
		};

	internal static bool? ParseNullableBool(string? cellValueText)
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
}
