namespace PanoramicData.SheetMagic;

/// <summary>
/// Chooses and orders the properties and extended keys written for each item
/// </summary>
internal static class ItemPropertySelector
{
	internal static (Type basicType, HashSet<string> keyHashSet) DetermineTypeAndCollectKeys<T>(
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

	internal static List<PropertyInfo> GetFilteredAndOrderedProperties(Type basicType, AddSheetOptions addSheetOptions)
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

	internal static List<string> SortKeyList(HashSet<string> keyHashSet, AddSheetOptions addSheetOptions)
		=> addSheetOptions.SortExtendedProperties
			? [.. keyHashSet.OrderBy(k => k)]
			: [.. keyHashSet];
}
