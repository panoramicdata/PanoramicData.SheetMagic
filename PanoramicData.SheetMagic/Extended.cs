using System.Collections.Generic;

namespace PanoramicData.SheetMagic;

public class Extended<T> where T : class
{
	public Extended(T? item, Dictionary<string, object?>? properties = null)
	{
		Item = item;
		Properties = properties ?? new Dictionary<string, object?>();
	}

	public T? Item { get; }

	public Dictionary<string, object?> Properties { get; }
}