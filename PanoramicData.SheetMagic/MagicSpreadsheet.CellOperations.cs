using System.Globalization;

namespace PanoramicData.SheetMagic;

/// <summary>
/// Cell creation and manipulation methods
/// </summary>
public partial class MagicSpreadsheet
{
	private static Cell CreateCell(string header, uint index, object? @object)
	{
		if (@object == null)
		{
			return CreateTextCell(header, index, string.Empty);
		}

		var objectTypeName = @object.GetType().Name;

		return objectTypeName switch
		{
			nameof(Int16) or nameof(Int32) or nameof(Int64) or
			nameof(UInt16) or nameof(UInt32) or nameof(UInt64)
				=> CreateNumericCell(header, index, Convert.ToDouble(@object)),

			nameof(Single) or nameof(Double) or nameof(Decimal)
				=> CreateNumericOrSpecialCell(header, index, @object),

			nameof(Boolean)
				=> CreateBooleanCell(header, index, (bool)@object),

			nameof(DateTime)
				=> CreateDateCell(header, index, (DateTime)@object),

			nameof(DateTimeOffset)
				=> CreateDateCell(header, index, ((DateTimeOffset)@object).UtcDateTime),

			_ when objectTypeName.StartsWith("Nullable`1")
				=> CreateNullableTypeCell(header, index, @object),

			_ => CreateTextCell(header, index, @object.ToString() ?? string.Empty),
		};
	}

	private static Cell CreateNumericOrSpecialCell(string header, uint index, object @object)
	{
		var doubleValue = Convert.ToDouble(@object);

		if (double.IsNaN(doubleValue))
		{
			return CreateTextCell(header, index, string.Empty);
		}

		if (double.IsPositiveInfinity(doubleValue))
		{
			return CreateTextCell(header, index, "Infinity");
		}

		if (double.IsNegativeInfinity(doubleValue))
		{
			return CreateTextCell(header, index, "-Infinity");
		}

		return CreateNumericCell(header, index, doubleValue);
	}

	private static Cell CreateNullableTypeCell(string header, uint index, object @object)
	{
		var objectTypeName = @object.GetType().Name;

		return objectTypeName switch
		{
			"Nullable`1<Single>" or "Nullable`1<Double>" or "Nullable`1<Decimal>"
				=> CreateNumericOrSpecialCell(header, index, @object),

			"Nullable`1<Boolean>"
				=> CreateBooleanCell(header, index, (bool)@object),

			"Nullable`1<DateTime>"
				=> CreateDateCell(header, index, (DateTime)@object),

			"Nullable`1<DateTimeOffset>"
				=> CreateDateCell(header, index, ((DateTimeOffset)@object).UtcDateTime),

			_ when objectTypeName.Contains("Int")
				=> CreateNumericCell(header, index, Convert.ToDouble(@object)),

			_ => CreateTextCell(header, index, @object.ToString() ?? string.Empty),
		};
	}

	private static Cell CreateNumericCell(string header, uint index, double number) =>
		new(new CellValue(number.ToString(CultureInfo.InvariantCulture)))
		{
			DataType = CellValues.Number,
			CellReference = header + index
		};

	private static Cell CreateDateCell(string header, uint index, DateTime date) =>
		new(new CellValue(date))
		{
			DataType = CellValues.Date,
			CellReference = header + index,
			StyleIndex = 1
		};

	private static Cell CreateTextCell(string header, uint index, string text) =>
		 new(new InlineString(new Text { Text = text }))
		 {
			 DataType = CellValues.InlineString,
			 CellReference = header + index
		 };

	private static Cell CreateBooleanCell(string header, uint index, bool booleanValue) =>
		new(new CellValue(booleanValue))
		{
			DataType = CellValues.Boolean,
			CellReference = header + index
		};

	private static string GetCellValueString(Cell cell, SharedStringTablePart? stringTable)
	{
		var cellValueText = cell.CellValue?.Text;

		if (cell.DataType != null && (CellValues)cell.DataType == CellValues.SharedString)
		{
			if (stringTable == null || cellValueText == null)
			{
				return string.Empty;
			}
			return stringTable.SharedStringTable.ElementAt(int.Parse(cellValueText)).InnerText;
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
		return (CellValues)cell.DataType! switch
		{
			CellValues.SharedString => GetSharedStringValue(stringTable, cellValueText),
			CellValues.Boolean => ParseBooleanValue(cellValueText),
			CellValues.Number => ParseNumberValue(cellValueText),
			CellValues.Date => DateTime.Parse(cellValueText!),
			CellValues.Error or CellValues.String or CellValues.InlineString => GetStringOrInfinityValue(cellValueText, cell),
			_ => throw new NotSupportedException($"Unsupported data type {cell.DataType?.Value.ToString() ?? "None"}"),
		};
	}

	private static string GetSharedStringValue(SharedStringTablePart? stringTable, string? cellValueText)
	{
		if (stringTable == null || cellValueText == null)
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
		if (cellValueText == "Infinity") return double.PositiveInfinity;
		if (cellValueText == "-Infinity") return double.NegativeInfinity;

		if (int.TryParse(cellValueText, out var doubleValue))
		{
			return doubleValue;
		}

		throw new FormatException($"Could not convert cell {cell.CellReference} to a double.");
	}

	private static float? ParseFloat(string? cellValueText, Cell cell)
	{
		// Handle special Infinity values
		if (cellValueText == "Infinity") return float.PositiveInfinity;
		if (cellValueText == "-Infinity") return float.NegativeInfinity;

		if (float.TryParse(cellValueText, out var floatValue))
		{
			return floatValue;
		}

		throw new FormatException($"Could not convert cell {cell.CellReference} to a float.");
	}

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
		return (cell.DataType is null ? null : (CellValues?)cell.DataType) switch
		{
			null => null,
			CellValues.SharedString => GetSharedStringValueTyped<T>(stringTable, cellValueText),
			CellValues.Boolean => ParseBooleanValueTyped(cellValueText),
			CellValues.Number => ParseNumberTyped(cellValueText),
			CellValues.Date => ParseDateTyped(cellValueText),
			CellValues.Error or CellValues.String or CellValues.InlineString => ParseStringOrInfinityTyped<T>(cellValueText, cell),
			_ => throw new NotSupportedException($"Unsupported data type {cell.DataType?.Value.ToString() ?? "None"}"),
		};
	}

	private static object? GetSharedStringValueTyped<T>(SharedStringTablePart? stringTable, string? cellValueText)
	{
		if (stringTable == null || cellValueText == null)
		{
			throw new FormatException("SharedStringTable or cell value text is null for SharedString type");
		}

		var stringTableIndex = int.Parse(cellValueText);
		var sharedStringElement = stringTable.SharedStringTable.ElementAt(stringTableIndex);
		var sharedStringValue = sharedStringElement.InnerText;

		// Handle special Infinity values for object-typed cells
		if (typeof(T).Name == "Object")
		{
			if (sharedStringValue == "Infinity") return double.PositiveInfinity;
			if (sharedStringValue == "-Infinity") return double.NegativeInfinity;
		}

		return sharedStringValue;
	}

	private static bool? ParseBooleanValueTyped(string? cellValueText)
		=> cellValueText switch
		{
			"1" or "true" => true,
			"0" or "false" => false,
			_ => null,
		};

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
		// For InlineString cells, get the actual text value
		var textValue = cellValueText;
		if (textValue == null && cell.DataType != null && cell.DataType == CellValues.InlineString)
		{
			// Extract text from InlineString element
			var inlineString = cell.Elements<InlineString>().FirstOrDefault();
			textValue = inlineString?.Text?.Text;
		}

		textValue ??= cell.InnerText;

		// Handle special Infinity values for object type
		if (typeof(T).Name == "Object")
		{
			if (textValue == "Infinity") return double.PositiveInfinity;
			if (textValue == "-Infinity") return double.NegativeInfinity;
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
}
