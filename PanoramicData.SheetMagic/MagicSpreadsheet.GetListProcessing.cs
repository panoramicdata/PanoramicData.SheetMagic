namespace PanoramicData.SheetMagic;

/// <summary>
/// Methods for reading sheets from the workbook - Part 2: Cell processing
/// </summary>
public partial class MagicSpreadsheet
{
	private void ProcessCellValue<T>(T item, Cell cell, System.Reflection.PropertyInfo property, string propertyName, SharedStringTablePart? stringTable) where T : class, new()
	{
		var propertyTypeName = property.PropertyType.IsGenericType
			 ? $"{property.PropertyType.GetGenericTypeDefinition().Name}<{string.Join(", ", property.PropertyType.GenericTypeArguments.Select(t => t.Name))}>"
			 : property.PropertyType.Name;

		if (TryProcessSimpleType(item, cell, property, propertyName, propertyTypeName, stringTable, _options))
		{
			return;
		}

		if (TryProcessNullableType(item, cell, propertyName, propertyTypeName, stringTable))
		{
			return;
		}

		// Handle enums and other types
		ProcessComplexType(item, cell, property, propertyName, stringTable);
	}

	private static bool TryProcessSimpleType<T>(
		T item,
		Cell cell,
		System.Reflection.PropertyInfo property,
		string propertyName,
		string propertyTypeName,
		SharedStringTablePart? stringTable,
		Options options) where T : class, new()
	{
		switch (propertyTypeName)
		{
			case "Double":
				ProcessDoubleValue(item, cell, propertyName, stringTable);
				return true;
			case "Single":
				ProcessSingleValue(item, cell, propertyName, stringTable);
				return true;
			case "Int16":
				ProcessInt16Value(item, cell, propertyName, stringTable);
				return true;
			case "UInt16":
				ProcessUInt16Value(item, cell, propertyName, stringTable);
				return true;
			case "Int32":
				ProcessInt32Value(item, cell, propertyName, stringTable);
				return true;
			case "UInt32":
				ProcessUInt32Value(item, cell, propertyName, stringTable);
				return true;
			case "Int64":
				ProcessInt64Value(item, cell, propertyName, stringTable);
				return true;
			case "UInt64":
				ProcessUInt64Value(item, cell, propertyName, stringTable);
				return true;
			case "Boolean":
				ProcessBooleanValue(item, cell, propertyName, stringTable);
				return true;
			case "DateTime":
				ProcessDateTimeValue(item, cell, propertyName, stringTable);
				return true;
			case "DateTimeOffset":
				ProcessDateTimeOffsetValue(item, cell, propertyName, stringTable);
				return true;
			case "String":
				SetItemProperty(item, (string?)GetCellValue<string>(cell, stringTable), propertyName);
				return true;
			case "List`1<String>":
				ProcessStringListValue(item, cell, propertyName, stringTable, options);
				return true;
			default:
				return false;
		}
	}

	private static bool TryProcessNullableType<T>(
		T item,
		Cell cell,
		string propertyName,
		string propertyTypeName,
		SharedStringTablePart? stringTable) where T : class, new()
	{
		switch (propertyTypeName)
		{
			case "Nullable`1<Boolean>":
				ProcessNullableBoolean(item, cell, propertyName, stringTable, propertyTypeName);
				return true;
			case "Nullable`1<Double>":
				ProcessNullableDouble(item, cell, propertyName, stringTable);
				return true;
			case "Nullable`1<Single>":
				ProcessNullableSingle(item, cell, propertyName, stringTable, propertyTypeName);
				return true;
			case "Nullable`1<Int64>":
				ProcessNullableInt64(item, cell, propertyName, stringTable, propertyTypeName);
				return true;
			case "Nullable`1<Int32>":
				ProcessNullableInt32(item, cell, propertyName, stringTable, propertyTypeName);
				return true;
			case "Nullable`1<Int16>":
				ProcessNullableInt16(item, cell, propertyName, stringTable, propertyTypeName);
				return true;
			case "Nullable`1<DateTime>":
				ProcessNullableDateTime(item, cell, propertyName, stringTable, propertyTypeName);
				return true;
			case "Nullable`1<DateTimeOffset>":
				ProcessNullableDateTimeOffset(item, cell, propertyName, stringTable, propertyTypeName);
				return true;
			default:
				return false;
		}
	}

	private static void ProcessComplexType<T>(
		T item,
		Cell cell,
		System.Reflection.PropertyInfo property,
		string propertyName,
		SharedStringTablePart? stringTable) where T : class, new()
	{
		var stringValue = (string?)GetCellValue<string>(cell, stringTable);
		if (property.PropertyType.IsEnum)
		{
			SetItemProperty(item, Enum.Parse(property.PropertyType, stringValue ?? string.Empty, true), propertyName);
		}
		else
		{
			throw new NotSupportedException($"Unsupported field type {property.PropertyType.Name}.");
		}
	}

	private static void ProcessDoubleValue<T>(T item, Cell cell, string propertyName, SharedStringTablePart? stringTable) where T : class, new()
	{
		var cellValueDoubleObject = GetCellValue<double>(cell, stringTable);
		if (cellValueDoubleObject != null)
		{
			SetItemProperty(item, Convert.ToDouble(cellValueDoubleObject), propertyName);
		}
	}

	private static void ProcessSingleValue<T>(T item, Cell cell, string propertyName, SharedStringTablePart? stringTable) where T : class, new()
	{
		var cellValueFloatObject = GetCellValue<float>(cell, stringTable);
		if (cellValueFloatObject != null)
		{
			SetItemProperty(item, Convert.ToSingle(cellValueFloatObject), propertyName);
		}
	}

	private static void ProcessInt16Value<T>(T item, Cell cell, string propertyName, SharedStringTablePart? stringTable) where T : class, new()
	{
		var cellValueShortObject = GetCellValue<short>(cell, stringTable);
		if (cellValueShortObject != null)
		{
			SetItemProperty(item, Convert.ToInt16(cellValueShortObject), propertyName);
		}
	}

	private static void ProcessUInt16Value<T>(T item, Cell cell, string propertyName, SharedStringTablePart? stringTable) where T : class, new()
	{
		var cellValueUShortObject = GetCellValue<ushort>(cell, stringTable);
		if (cellValueUShortObject != null)
		{
			SetItemProperty(item, Convert.ToInt16(cellValueUShortObject), propertyName);
		}
	}

	private static void ProcessInt32Value<T>(T item, Cell cell, string propertyName, SharedStringTablePart? stringTable) where T : class, new()
	{
		var cellValueIntObject = GetCellValue<int>(cell, stringTable);
		if (cellValueIntObject != null)
		{
			SetItemProperty(item, Convert.ToInt32(cellValueIntObject), propertyName);
		}
	}

	private static void ProcessUInt32Value<T>(T item, Cell cell, string propertyName, SharedStringTablePart? stringTable) where T : class, new()
	{
		var cellValueUIntObject = GetCellValue<uint>(cell, stringTable);
		if (cellValueUIntObject != null)
		{
			SetItemProperty(item, Convert.ToUInt32(cellValueUIntObject), propertyName);
		}
	}

	private static void ProcessInt64Value<T>(T item, Cell cell, string propertyName, SharedStringTablePart? stringTable) where T : class, new()
	{
		var cellValueLongObject = GetCellValue<long>(cell, stringTable);
		if (cellValueLongObject != null)
		{
			SetItemProperty(item, Convert.ToInt64(cellValueLongObject), propertyName);
		}
	}

	private static void ProcessUInt64Value<T>(T item, Cell cell, string propertyName, SharedStringTablePart? stringTable) where T : class, new()
	{
		var cellValueULongObject = GetCellValue<ulong>(cell, stringTable);
		if (cellValueULongObject != null)
		{
			SetItemProperty(item, Convert.ToUInt64(cellValueULongObject), propertyName);
		}
	}

	private static void ProcessBooleanValue<T>(T item, Cell cell, string propertyName, SharedStringTablePart? stringTable) where T : class, new()
	{
		var cellValue = GetCellValue<bool>(cell, stringTable);
		if (cellValue != null)
		{
			SetItemProperty(item, cellValue, propertyName);
		}
	}

	private static void ProcessDateTimeValue<T>(T item, Cell cell, string propertyName, SharedStringTablePart? stringTable) where T : class, new()
	{
		var cellValueDateTimeObject = GetCellValue<DateTime>(cell, stringTable);
		if (cellValueDateTimeObject != null)
		{
			SetItemProperty(item, Convert.ToDateTime(cellValueDateTimeObject), propertyName);
		}
	}

	private static void ProcessDateTimeOffsetValue<T>(T item, Cell cell, string propertyName, SharedStringTablePart? stringTable) where T : class, new()
	{
		var cellValueDateTimeOffsetObject = GetCellValue<DateTimeOffset>(cell, stringTable);
		if (cellValueDateTimeOffsetObject != null)
		{
			SetItemProperty(item, Convert.ToDateTime(cellValueDateTimeOffsetObject), propertyName);
		}
	}

	private static void ProcessStringListValue<T>(T item, Cell cell, string propertyName, SharedStringTablePart? stringTable, Options options) where T : class, new()
	{
		var text = (string?)GetCellValue<string>(cell, stringTable);
		if (text is null)
		{
			SetItemProperty(item, new List<string>(), propertyName);
		}
		else
		{
			var stringList = text.Split([options.ListSeparator], StringSplitOptions.RemoveEmptyEntries).ToList();
			SetItemProperty(item, stringList, propertyName);
		}
	}

	private static void ProcessNullableBoolean<T>(T item, Cell cell, string propertyName, SharedStringTablePart? stringTable, string propertyTypeName) where T : class, new()
	{
		var cellValue = GetCellValue<object?>(cell, stringTable);
		switch (cellValue)
		{
			case string stringValue2:
				if (string.IsNullOrWhiteSpace(stringValue2))
				{
					SetItemProperty(item, (bool?)null, propertyName);
				}
				else
				{
					SetItemProperty(item, (bool?)bool.Parse(stringValue2), propertyName);
				}
				break;
			case bool typedValue2:
				SetItemProperty(item, (bool?)typedValue2, propertyName);
				break;
			case null:
				SetItemProperty(item, null, propertyName);
				break;
			default:
				throw new InvalidOperationException($"Invalid {propertyTypeName} value type for {cellValue}: {cellValue.GetType().Name}");
		}
	}

	private static void ProcessNullableDouble<T>(T item, Cell cell, string propertyName, SharedStringTablePart? stringTable) where T : class, new()
	{
		var cellValue = GetCellValue<object?>(cell, stringTable);
		switch (cellValue)
		{
			case string stringValue2:
				if (string.IsNullOrWhiteSpace(stringValue2))
				{
					SetItemProperty(item, (double?)null, propertyName);
				}
				else
				{
					SetItemProperty(item, (double?)double.Parse(stringValue2), propertyName);
				}
				break;
			case double typedValue2:
				SetItemProperty(item, (double?)typedValue2, propertyName);
				break;
			case null:
				SetItemProperty(item, null, propertyName);
				break;
			default:
				throw new InvalidOperationException("Invalid value type");
		}
	}

	private static void ProcessNullableSingle<T>(T item, Cell cell, string propertyName, SharedStringTablePart? stringTable, string propertyTypeName) where T : class, new()
	{
		var cellValue = GetCellValue<object?>(cell, stringTable);
		switch (cellValue)
		{
			case string stringValue2:
				if (string.IsNullOrWhiteSpace(stringValue2))
				{
					SetItemProperty(item, (float?)null, propertyName);
				}
				else
				{
					SetItemProperty(item, (float?)float.Parse(stringValue2), propertyName);
				}
				break;
			case float typedValue2:
				SetItemProperty(item, (float?)typedValue2, propertyName);
				break;
			case null:
				SetItemProperty(item, null, propertyName);
				break;
			default:
				throw new InvalidOperationException($"Invalid {propertyTypeName} value type for {cellValue}: {cellValue.GetType().Name}");
		}
	}

	private static void ProcessNullableInt64<T>(T item, Cell cell, string propertyName, SharedStringTablePart? stringTable, string propertyTypeName) where T : class, new()
	{
		var cellValue = GetCellValue<object?>(cell, stringTable);
		switch (cellValue)
		{
			case string stringValue2:
				if (string.IsNullOrWhiteSpace(stringValue2))
				{
					SetItemProperty(item, (long?)null, propertyName);
				}
				else
				{
					SetItemProperty(item, (long?)long.Parse(stringValue2), propertyName);
				}
				break;
			case long typedValue2:
				SetItemProperty(item, (long?)typedValue2, propertyName);
				break;
			case null:
				SetItemProperty(item, null, propertyName);
				break;
			default:
				throw new InvalidOperationException($"Invalid {propertyTypeName} value type for {cellValue}: {cellValue.GetType().Name}");
		}
	}

	private static void ProcessNullableInt32<T>(T item, Cell cell, string propertyName, SharedStringTablePart? stringTable, string propertyTypeName) where T : class, new()
	{
		var cellValue = GetCellValue<object?>(cell, stringTable);
		switch (cellValue)
		{
			case string stringValue2:
				if (string.IsNullOrWhiteSpace(stringValue2))
				{
					SetItemProperty(item, (int?)null, propertyName);
				}
				else
				{
					SetItemProperty(item, (int?)int.Parse(stringValue2), propertyName);
				}
				break;
			case int typedValue2:
				SetItemProperty(item, (int?)typedValue2, propertyName);
				break;
			case double typedValue2:
				SetItemProperty(item, (int?)typedValue2, propertyName);
				break;
			case null:
				SetItemProperty(item, null, propertyName);
				break;
			default:
				throw new InvalidOperationException($"Invalid {propertyTypeName} value type for {cellValue}: {cellValue.GetType().Name}");
		}
	}

	private static void ProcessNullableInt16<T>(T item, Cell cell, string propertyName, SharedStringTablePart? stringTable, string propertyTypeName) where T : class, new()
	{
		var cellValue = GetCellValue<object?>(cell, stringTable);
		switch (cellValue)
		{
			case string stringValue2:
				if (string.IsNullOrWhiteSpace(stringValue2))
				{
					SetItemProperty(item, (short?)null, propertyName);
				}
				else
				{
					SetItemProperty(item, (short?)short.Parse(stringValue2), propertyName);
				}
				break;
			case short typedValue2:
				SetItemProperty(item, (short?)typedValue2, propertyName);
				break;
			case null:
				SetItemProperty(item, null, propertyName);
				break;
			default:
				throw new InvalidOperationException($"Invalid {propertyTypeName} value type for {cellValue}: {cellValue.GetType().Name}");
		}
	}

	private static void ProcessNullableDateTime<T>(T item, Cell cell, string propertyName, SharedStringTablePart? stringTable, string propertyTypeName) where T : class, new()
	{
		var cellValue = GetCellValue<object?>(cell, stringTable);
		switch (cellValue)
		{
			case string stringValue2:
				if (string.IsNullOrWhiteSpace(stringValue2))
				{
					SetItemProperty(item, (DateTime?)null, propertyName);
				}
				else
				{
					SetItemProperty(item, (DateTime?)DateTime.Parse(stringValue2), propertyName);
				}
				break;
			case DateTime typedValue2:
				SetItemProperty(item, (DateTime?)typedValue2, propertyName);
				break;
			case null:
				SetItemProperty(item, null, propertyName);
				break;
			default:
				throw new InvalidOperationException($"Invalid {propertyTypeName} value type for {cellValue}: {cellValue.GetType().Name}");
		}
	}

	private static void ProcessNullableDateTimeOffset<T>(T item, Cell cell, string propertyName, SharedStringTablePart? stringTable, string propertyTypeName) where T : class, new()
	{
		var cellValue = GetCellValue<object?>(cell, stringTable);
		switch (cellValue)
		{
			case string stringValue2:
				if (string.IsNullOrWhiteSpace(stringValue2))
				{
					SetItemProperty(item, (DateTimeOffset?)null, propertyName);
				}
				else
				{
					SetItemProperty(item, (DateTimeOffset?)DateTimeOffset.Parse(stringValue2), propertyName);
				}
				break;
			case DateTime typedValue2:
				SetItemProperty(item, new DateTimeOffset(typedValue2, TimeSpan.Zero), propertyName);
				break;
			case DateTimeOffset typedValue2:
				SetItemProperty(item, (DateTimeOffset?)typedValue2, propertyName);
				break;
			case null:
				SetItemProperty(item, null, propertyName);
				break;
			default:
				throw new InvalidOperationException($"Invalid {propertyTypeName} value type for {cellValue}: {cellValue.GetType().Name}");
		}
	}
}
