using System.Collections.Generic;

namespace PanoramicData.SheetMagic
{
	/// <summary>
	/// Excel's built-in formats. See: http://polymathprogrammer.com/2011/02/15/built-in-styles-for-excel-open-xml/
	/// </summary>
	public static class BuiltInCellFormats
	{
		private static readonly Dictionary<int, string> _cellFormatsDictionary = new Dictionary<int, string>()
		{
			{ 0, "" },
			{ 1, "0" },
			{ 2, "0.00" },
			{ 3, "#,##0" },
			{ 4, "#,##0.00" },
			{ 9, "0%" },
			{ 10, "0.00%" },
			{ 11, "0.00E+00" },
			{ 12, "# ?/?" },
			{ 13, "# ??/??" },
			{ 14, "mm-dd-yy" },
			{ 15, "d-mmm-yy" },
			{ 16, "d-mmm" },
			{ 17, "mmm-yy" },
			{ 18, "h:mm AM/PM" },
			{ 19, "h:mm:ss AM/PM" },
			{ 20, "h:mm" },
			{ 21, "h:mm:ss" },
			{ 22, "m/d/yy h:mm" },
			{ 37, "#,##0 ;(#,##0)" },
			{ 38, "#,##0 ;[Red](#,##0)" },
			{ 39, "#,##0.00;(#,##0.00)" },
			{ 40, "#,##0.00;[Red](#,##0.00)" },
			{ 45, "mm:ss" },
			{ 46, "[h]:mm:ss" },
			{ 47, "mmss.0" },
			{ 48, "##0.0E+0" },
			{ 49, "@" }
		};


		public static string? GetBuiltInCellFormatByIndex(int styleIndex)
			=> _cellFormatsDictionary.ContainsKey(styleIndex)
				? _cellFormatsDictionary[styleIndex]
				: null;
	}
}
