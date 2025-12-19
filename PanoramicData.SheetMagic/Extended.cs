namespace PanoramicData.SheetMagic;

/// <summary>
/// Wraps an item of type T with additional extended properties that don't map to the type's properties.
/// Used for reading/writing data with extra columns that aren't defined in the model.
/// </summary>
/// <typeparam name="T">The type of the main item.</typeparam>
public class Extended<T> where T : class
{
	/// <summary>
	/// Creates a new Extended instance with an item and properties.
	/// </summary>
	/// <param name="item">The main item.</param>
	/// <param name="properties">Dictionary of extended properties.</param>
	public Extended(T? item, Dictionary<string, object?> properties)
	{
		Item = item;
		Properties = properties;
	}

	/// <summary>
	/// Creates a new Extended instance with just an item and empty properties.
	/// </summary>
	/// <param name="item">The main item.</param>
	public Extended(T? item)
	{
		Item = item;
		Properties = [];
	}

	/// <summary>
	/// Gets the main item.
	/// </summary>
	public T? Item { get; }

	/// <summary>
	/// Gets the dictionary of extended properties that don't map to the item's type.
	/// </summary>
	public Dictionary<string, object?> Properties { get; }
}