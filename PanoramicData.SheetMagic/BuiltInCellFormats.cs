namespace PanoramicData.SheetMagic;

/// <summary>
/// Specifies the type of cell format.
/// </summary>
public enum CellFormatType
{
	/// <summary>Unknown or unrecognized format.</summary>
	Unknown = 0,
	/// <summary>General format.</summary>
	General,
	/// <summary>Text format.</summary>
	Text,
	/// <summary>Numeric format.</summary>
	Number,
	/// <summary>Date and/or time format.</summary>
	DateTime
}

/// <summary>
/// Excel's built-in formats. See: http://polymathprogrammer.com/2011/02/15/built-in-styles-for-excel-open-xml/
/// </summary>
public static class BuiltInCellFormats
{
	internal const int CustomFormatStartIndex = 164;

	private static readonly Dictionary<int, (string formatString, CellFormatType formatType)> _cellFormatsDictionary
		= new()
	{
		// Some of these may not work when doing a number.ToString() but can tweak over time
		// Negative ones with colours have been updated i.e. removed the second part and colour like #;[Red](#)
		// Some of these are purely returned as text
		{ 0, ("", CellFormatType.General) },
		{ 1, ("0", CellFormatType.Number) },
		{ 2, ("0.00", CellFormatType.Number) },
		{ 3, ("#,##0", CellFormatType.Number) },
		{ 4, ("#,##0.00", CellFormatType.Number) },
		{ 9, ("0%", CellFormatType.Number) },
		{ 10, ("0.00%", CellFormatType.Number) },
		{ 11, ("0.00E+00", CellFormatType.Text) },			// Scientific - return as string
		{ 12, ("# ?/?", CellFormatType.Text) },				// Fractions - return as string
		{ 13, ("# ??/??", CellFormatType.Text) },			// Fractions - return as string
		{ 14, ("dd/mm/yyyy", CellFormatType.DateTime) },
		{ 15, ("d/mmm/yy", CellFormatType.DateTime) },
		{ 16, ("d/mmm", CellFormatType.DateTime) },
		{ 17, ("mmm/yy", CellFormatType.DateTime) },
		{ 18, ("h:mm AM/PM", CellFormatType.DateTime) },
		{ 19, ("h:mm:ss AM/PM", CellFormatType.DateTime) },
		{ 20, ("h:mm", CellFormatType.DateTime) },
		{ 21, ("h:mm:ss", CellFormatType.DateTime) },
		{ 22, ("m/d/yy h:mm", CellFormatType.DateTime) },
		{ 37, ("#,##0", CellFormatType.Number) },
		{ 38, ("#,##0", CellFormatType.Number) },
		{ 39, ("#,##0.00", CellFormatType.Number) },
		{ 40, ("#,##0.00", CellFormatType.Number) },
		{ 45, ("mm:ss", CellFormatType.DateTime) },
		{ 46, ("[h]:mm:ss", CellFormatType.DateTime) },
		{ 47, ("mmss.0", CellFormatType.DateTime) },
		{ 48, ("##0.0E+0", CellFormatType.Text) },
		{ 49, ("@", CellFormatType.Text) }
	};

	/// <summary>
	/// Gets the built-in cell format information by its index.
	/// </summary>
	/// <param name="styleIndex">The style index to look up.</param>
	/// <returns>A tuple containing the format string and format type, or null if not found.</returns>
	public static (string formatString, CellFormatType formatType)? GetBuiltInCellFormatByIndex(int styleIndex)
		=> _cellFormatsDictionary.TryGetValue(styleIndex, out var value)
			? value
			: null;
}