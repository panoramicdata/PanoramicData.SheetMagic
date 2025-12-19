namespace PanoramicData.SheetMagic;

/// <summary>
/// Property reflection helper methods
/// </summary>
public partial class MagicSpreadsheet
{
	private static PropertyInfo GetPropertyInfo(string path, IEnumerable<PropertyInfo> props)
	{
		if (string.IsNullOrWhiteSpace(path))
		{
			throw new ArgumentException("Property path must be specified", nameof(path));
		}

		// Nested path?
		if (path.Contains('.'))
		{
			var parentPropertyName = path[..path.IndexOf('.')];
			var parentProp = props.FirstOrDefault(x => x.Name.Equals(parentPropertyName, StringComparison.InvariantCultureIgnoreCase))
				?? throw new PropertyNotFoundException(parentPropertyName);

			try
			{
				// Recurse path
				var parentTypeProps = parentProp.PropertyType.GetProperties();
				return GetPropertyInfo(path[(parentPropertyName.Length + 1)..], parentTypeProps);
			}
			catch (PropertyNotFoundException)
			{
				throw new PropertyNotFoundException(path);
			}
		}
		else
		{
			var p = props.FirstOrDefault(x => x.Name.Equals(path, StringComparison.InvariantCultureIgnoreCase));
			return p is null
				? throw new PropertyNotFoundException(path)
				: p;
		}
	}

	private static object? GetPropertyValue(string path, object? item)
	{
		if (item is null)
		{
			return null;
		}

		var props = item.GetType().GetProperties();

		// Nested path?
		if (path.Contains('.'))
		{
			var parentPropertyName = path[..path.IndexOf('.')];
			var parentProp = props.FirstOrDefault(x => x.Name.Equals(parentPropertyName, StringComparison.InvariantCultureIgnoreCase))
				?? throw new PropertyNotFoundException(parentPropertyName);

			try
			{
				// Recurse path
				var parentItem = parentProp.GetValue(item);
				return GetPropertyValue(path[(parentPropertyName.Length + 1)..], parentItem);
			}
			catch (PropertyNotFoundException)
			{
				throw new PropertyNotFoundException(path);
			}
		}
		else
		{
			var prop = props.FirstOrDefault(x => x.Name.Equals(path, StringComparison.InvariantCultureIgnoreCase));
			return prop?.GetValue(item);
		}
	}

	private static void SetItemProperty<T, T1>(T item, T1 cellValue, string propertyName)
	{
		var cellValues = new List<object?> { cellValue };
		_ = typeof(T).InvokeMember(propertyName,
			 BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty,
			 Type.DefaultBinder, item, [.. cellValues]);
	}

	private static void SetItemProperty<T>(T item, object? cellValue, string propertyName)
	{
		var cellValues = new List<object?> { cellValue };
		_ = typeof(T).InvokeMember(propertyName,
			 BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty,
			 Type.DefaultBinder, item, [.. cellValues]);
	}
}
