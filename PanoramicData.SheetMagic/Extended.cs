using System.Collections.Generic;

namespace PanoramicData.SheetMagic;

public class Extended<T> where T : class
{
	public Extended(T? item, Dictionary<string, object?> properties)
	{
		Item = item;
		Properties = properties;
	}

	public Extended(T? item)
	{
		Item = item;
		Properties = new Dictionary<string, object?>();
	}

	public T? Item { get; }

	public Dictionary<string, object?> Properties { get; }
}