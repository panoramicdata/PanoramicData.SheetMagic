using System.ComponentModel;

namespace PanoramicData.SheetMagic.Extensions;

public static class Attributes
{
	public static string? GetPropertyDescription(this PropertyInfo propertyInfo)
		=> propertyInfo.GetCustomAttributes(typeof(DescriptionAttribute)) is not DescriptionAttribute[] descriptions || descriptions.Length == 0
			? null
			: descriptions[0].Description;
}
