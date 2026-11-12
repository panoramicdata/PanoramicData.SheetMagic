namespace PanoramicData.SheetMagic;

/// <summary>
/// Typed cell value reading for cells that carry a data type
/// </summary>
internal static class TypedDataTypeCellReader
{
	internal static object? GetCellValueWithDataType<T>(Cell cell, SharedStringTablePart? stringTable, string? cellValueText)
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

	internal static object? GetTextualCellValueWithDataType<T>(Cell cell, CellValues dataType, string? cellValueText)
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

	internal static object? GetSharedStringValueTyped<T>(SharedStringTablePart? stringTable, string? cellValueText)
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
	internal static double? TryParseInfinityForObject<T>(string? text)
		=> typeof(T).Name == "Object"
			? TryParseInfinity(text)
			: null;
}
