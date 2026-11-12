using PanoramicData.SheetMagic.Extensions;

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
}
