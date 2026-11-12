using System.Collections;
using System.Text;

namespace PanoramicData.SheetMagic;

/// <summary>
/// Row and cell writing for items
/// </summary>
public partial class MagicSpreadsheet
{
	private void AddDataRows<T>(
		List<T> items,
		AddSheetOptions addSheetOptions,
		Type type,
		bool isExtended,
		SheetData sheetData,
		List<PropertyInfo> propertyList,
		List<string> keyList)
	{
		var enumerableCellOptions = addSheetOptions?.EnumerableCellOptions ?? _options.EnumerableCellOptions;
		uint rowIndex = 1;

		foreach (var item in items)
		{
			var row = new Row { RowIndex = ++rowIndex };
			_ = sheetData.AppendChild(row);
			var writer = new RowWriter(row, rowIndex);

			AddItemCells(item, addSheetOptions!, type, isExtended, propertyList, enumerableCellOptions, writer);

			if (isExtended)
			{
				AddExtendedPropertyCells(item, type, keyList, writer);
			}
		}
	}

	private static void AddItemCells<T>(
		T item,
		AddSheetOptions addSheetOptions,
		Type type,
		bool isExtended,
		List<PropertyInfo> propertyList,
		EnumerableCellOptions enumerableCellOptions,
		RowWriter writer)
	{
		if (addSheetOptions?.PropertyOrder?.Length > 0)
		{
			AddOrderedPropertyCells(item, addSheetOptions, enumerableCellOptions, writer);
		}
		else
		{
			AddStandardPropertyCells(item, type, isExtended, propertyList, enumerableCellOptions, writer);
		}
	}

	private static void AddOrderedPropertyCells<T>(
		T item,
		AddSheetOptions addSheetOptions,
		EnumerableCellOptions enumerableCellOptions,
		RowWriter writer)
	{
		foreach (var prop in addSheetOptions!.PropertyOrder!)
		{
			writer.AppendValue(ConvertValueForCell(enumerableCellOptions, GetPropertyValue(prop, item)));
		}
	}

	private static void AddStandardPropertyCells<T>(
		T item,
		Type type,
		bool isExtended,
		List<PropertyInfo> propertyList,
		EnumerableCellOptions enumerableCellOptions,
		RowWriter writer)
	{
		foreach (var property in propertyList)
		{
			var propertyValue = GetPropertyValueForCell(item, type, isExtended, property);
			writer.AppendValue(ConvertValueForCell(enumerableCellOptions, propertyValue));
		}
	}

	private static object? GetPropertyValueForCell<T>(T item, Type type, bool isExtended, PropertyInfo property)
	{
		if (isExtended)
		{
			var baseItem = type.GetProperties().Single(p => p.Name == nameof(Extended<object>.Item)).GetValue(item);
			return property.GetValue(baseItem);
		}
		else
		{
			return property.GetValue(item);
		}
	}

	private static void AddExtendedPropertyCells<T>(
		T item,
		Type type,
		List<string> keyList,
		RowWriter writer)
	{
		var propertyInfo = type.GetProperties().Single(p => p.Name == nameof(Extended<object>.Properties));
		var dictionary = (Dictionary<string, object>?)propertyInfo.GetValue(item);

		if (dictionary == null)
		{
			return;
		}

		foreach (var key in keyList)
		{
			if (!dictionary.TryGetValue(key, out var @object))
			{
				@object = string.Empty;
			}

			// Don't add cells for null objects
			if (@object is null)
			{
				writer.SkipCell();
				continue;
			}

			writer.AppendValue(@object);
		}
	}

	/// <summary>
	/// Writes cells left to right into a single row, keeping track of the column each one lands in.
	/// </summary>
	private sealed class RowWriter(Row row, uint rowIndex)
	{
		private int _cellIndex;

		/// <summary>
		/// Writes <paramref name="value"/> into the next column.
		/// </summary>
		internal void AppendValue(object? value)
		{
			_ = row.AppendChild(CreateCell(ColumnLetter(_cellIndex), rowIndex, value));
			_cellIndex++;
		}

		/// <summary>
		/// Leaves the next column empty.
		/// </summary>
		internal void SkipCell() => _cellIndex++;
	}

	private static object? ConvertValueForCell<T>(EnumerableCellOptions enumerableCellOptions, T? v)
	{
		if (enumerableCellOptions.Expand && v is not string && v is IEnumerable iEnumerable)
		{
			return ExpandEnumerableToString(iEnumerable, enumerableCellOptions.CellDelimiter ?? ", ");
		}

		return v is not null ? v is string ? v.ToString() : v : string.Empty;
	}

	private static string ExpandEnumerableToString(IEnumerable iEnumerable, string delimiter)
	{
		var stringBuilder = new StringBuilder();
		var isFirst = true;

		foreach (var il in iEnumerable)
		{
			if (!isFirst)
			{
				_ = stringBuilder.Append(delimiter);
			}

			_ = stringBuilder.Append(il?.ToString() ?? "NULL");
			isFirst = false;
		}

		return stringBuilder.ToString();
	}
}
