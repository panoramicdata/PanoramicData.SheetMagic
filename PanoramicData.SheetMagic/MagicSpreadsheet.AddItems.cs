using PanoramicData.SheetMagic.Extensions;
using System.Collections;
using System.Text;

namespace PanoramicData.SheetMagic;

/// <summary>
/// Methods for processing and adding items to sheets
/// </summary>
public partial class MagicSpreadsheet
{
	private SheetLayout AddItems<T>(
		List<T> items,
		AddSheetOptions addSheetOptions,
		Type type,
		bool isExtended,
		SheetData sheetData)
	{
		var (basicType, keyHashSet) = DetermineTypeAndCollectKeys(items, addSheetOptions, type, isExtended);

		var propertyList = GetFilteredAndOrderedProperties(basicType, addSheetOptions);
		var keyList = SortKeyList(keyHashSet, addSheetOptions);

		var totalColumnCount = (uint)(propertyList.Count + keyList.Count);
		var columnConfigurations = CreateColumns(totalColumnCount);

		AddHeaderRow(sheetData, propertyList, keyList, addSheetOptions);
		AddDataRows(items, addSheetOptions, type, isExtended, sheetData, propertyList, keyList);

		return new SheetLayout(propertyList, columnConfigurations, keyList, totalColumnCount);
	}

	private static (Type basicType, HashSet<string> keyHashSet) DetermineTypeAndCollectKeys<T>(
		List<T> items,
		AddSheetOptions addSheetOptions,
		Type type,
		bool isExtended)
	{
		var keyHashSet = new HashSet<string>();
		Type basicType;

		if (isExtended)
		{
			basicType = type.GenericTypeArguments[0];
			CollectExtendedPropertyKeys(items, addSheetOptions, type, keyHashSet);
		}
		else
		{
			basicType = type;
		}

		return (basicType, keyHashSet);
	}

	private static void CollectExtendedPropertyKeys<T>(
		List<T> items,
		AddSheetOptions addSheetOptions,
		Type type,
		HashSet<string> keyHashSet)
	{
		var propertyInfo = type.GetProperties().Single(p => p.Name == nameof(Extended<object>.Properties));

		foreach (var item in items)
		{
			if (item is null)
			{
				continue;
			}

			var dictionary = (Dictionary<string, object>?)propertyInfo.GetValue(item);
			if (dictionary == null)
			{
				continue;
			}

			var keys = FilterKeys(dictionary.Keys.ToList(), addSheetOptions);

			foreach (var key in keys)
			{
				_ = keyHashSet.Add(key);
			}
		}
	}

	private static List<string> FilterKeys(List<string> keys, AddSheetOptions addSheetOptions)
	{
		if (addSheetOptions.IncludeProperties?.Count > 0)
		{
			return [.. keys.Where(key => addSheetOptions.IncludeProperties.Contains(key, StringComparer.InvariantCultureIgnoreCase))];
		}

		if (addSheetOptions.ExcludeProperties?.Count > 0)
		{
			return [.. keys.Where(key => !addSheetOptions.ExcludeProperties.Contains(key, StringComparer.InvariantCultureIgnoreCase))];
		}

		return keys;
	}

	private static List<PropertyInfo> GetFilteredAndOrderedProperties(Type basicType, AddSheetOptions addSheetOptions)
	{
		var propertyList = new List<PropertyInfo>();
		propertyList.AddRange(basicType.GetProperties());

		propertyList = FilterProperties(propertyList, addSheetOptions);

		if (addSheetOptions.PropertyOrder?.Length > 0)
		{
			propertyList = OrderProperties(propertyList, addSheetOptions.PropertyOrder);
		}

		return propertyList;
	}

	private static List<PropertyInfo> FilterProperties(List<PropertyInfo> propertyList, AddSheetOptions addSheetOptions)
	{
		if (addSheetOptions.IncludeProperties?.Count > 0)
		{
			return [.. propertyList.Where(p => addSheetOptions.IncludeProperties.Contains(p.Name, StringComparer.InvariantCultureIgnoreCase))];
		}

		if (addSheetOptions.ExcludeProperties?.Count > 0)
		{
			return [.. propertyList.Where(p => !addSheetOptions.ExcludeProperties.Contains(p.Name, StringComparer.InvariantCultureIgnoreCase))];
		}

		return propertyList;
	}

	private static List<PropertyInfo> OrderProperties(List<PropertyInfo> propertyList, string[] propertyOrder)
	{
		var orderedPropertyList = new List<PropertyInfo>();
		foreach (var prop in propertyOrder)
		{
			// Support nested properties in form: x.y or x.y.z etc
			var p = GetPropertyInfo(prop, propertyList);
			if (p != null)
			{
				orderedPropertyList.Add(p);
			}
		}

		return orderedPropertyList;
	}

	private static List<string> SortKeyList(HashSet<string> keyHashSet, AddSheetOptions addSheetOptions)
		=> addSheetOptions.SortExtendedProperties
			? [.. keyHashSet.OrderBy(k => k)]
			: [.. keyHashSet];

	private static Columns CreateColumns(uint totalColumnCount)
	{
		var columnConfigurations = new Columns();
		for (var n = 0; n < totalColumnCount; n++)
		{
			_ = columnConfigurations.AppendChild(new Column { BestFit = true });
		}
		return columnConfigurations;
	}

	private static void AddHeaderRow(
		SheetData sheetData,
		List<PropertyInfo> propertyList,
		List<string> keyList,
		AddSheetOptions addSheetOptions)
	{
		uint rowIndex = 1;
		var row = new Row { RowIndex = rowIndex };
		_ = sheetData.AppendChild(row);
		var writer = new RowWriter(row, rowIndex);

		var headers = GetHeaders(propertyList, addSheetOptions);
		AddHeaderCells(writer, headers);
		AddHeaderCells(writer, keyList);
	}

	private static string[] GetHeaders(List<PropertyInfo> propertyList, AddSheetOptions addSheetOptions)
		=> addSheetOptions.PropertyHeaders?.Length == propertyList.Count
			? addSheetOptions.PropertyHeaders
			: [.. propertyList.Select(p => p.GetPropertyDescription() ?? p.Name)];

	private static void AddHeaderCells(RowWriter writer, IEnumerable<string> headers)
	{
		foreach (var header in headers)
		{
			writer.AppendValue(header ?? string.Empty);
		}
	}

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
