using System.ComponentModel;

namespace PanoramicData.SheetMagic.Extensions;

/// <summary>
/// Extension methods for working with property attributes.
/// </summary>
public static class Attributes
{
	/// <summary>
	/// Gets the description from a DescriptionAttribute on a property, if present.
	/// </summary>
	/// <param name="propertyInfo">The property to get the description from.</param>
	/// <returns>The description string, or null if no DescriptionAttribute is present.</returns>
	public static string? GetPropertyDescription(this PropertyInfo propertyInfo)
		=> propertyInfo.GetCustomAttributes<DescriptionAttribute>() is not DescriptionAttribute[] descriptions || descriptions.Length == 0
			? null
			: descriptions[0].Description;
}
