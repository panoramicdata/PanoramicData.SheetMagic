namespace PanoramicData.SheetMagic;

/// <summary>
/// Cell formatting methods
/// </summary>
public partial class MagicSpreadsheet
{
	private static string? FormatCellAsNumber(Cell cell, string formatString)
		// Check if it's a number
		=> double.TryParse(
			string.IsNullOrEmpty(cell.CellValue?.Text)
				? cell.InnerText
				: cell.CellValue!.Text,
			out var number
			)
				? number.ToString(formatString)
				: null;

	private static string? FormatCellAsDateTime(Cell cell, string formatString)
	{
		// Excel stores dates as a number (number of days since January 1, 1900),
		//so "44166" text is 03/12/2020
		// IF the number is an integer, it's only days. If it's a double, it's a fractional
		// portion of a day.
		var baseDate = new DateTime(1900, 01, 01);

		if (int.TryParse(
			(cell.CellValue != null &&
			!string.IsNullOrEmpty(cell.CellValue.Text))
			? cell.CellValue.Text
			: cell.InnerText, out var intDaysSinceBaseDate))
		{
			// See: https://www.kirix.com/stratablog/excel-date-conversion-days-from-1900
			// Note you DO have to take off 2 days!
			DateTime? actualDate = baseDate.AddDays(intDaysSinceBaseDate).AddDays(-2);

			// Return the date - we have to replace lower-case 'm' with upper-case as
			// required by C# else we get minutes
			// Some custom formats used by customers also have @ and ; in them.
			// Use InvariantCulture so the '/' date separator is always a literal slash.
			return actualDate.Value.ToString(
				formatString
					.Replace("\\", string.Empty)
					.Replace(";", string.Empty)
					.Replace("@", string.Empty)
					.Replace("m", "M"),
				System.Globalization.CultureInfo.InvariantCulture)
				.Trim();
		}

		// Could not parse cell value as an integer
		return null;
	}

	/// <summary>
	/// Is the format string a date string?
	/// </summary>
	/// <param name="formatString"></param>
	/// <returns></returns>
	private static bool IsFormatStringADate(string formatString) =>
		formatString.Contains('d', StringComparison.OrdinalIgnoreCase) ||
		formatString.Contains('m', StringComparison.OrdinalIgnoreCase) ||
		formatString.Contains('y', StringComparison.OrdinalIgnoreCase);

	private string? GetCellFormatFromStyle(Cell cell)
	{
		try
		{
			if (!HasStyleIndex(cell))
			{
				return null;
			}

			var styleIndex = (int)cell.StyleIndex!.Value;
			var (cellFormats, numberingFormats) = GetFormattingParts();

			if (cellFormats == null)
			{
				return null;
			}

			var cellFormat = (CellFormat)cellFormats.ElementAt(styleIndex);

			if (cellFormat.NumberFormatId?.HasValue != true)
			{
				return null;
			}

			var formatString = GetFormatString(cellFormat.NumberFormatId.Value, numberingFormats);

			if (string.IsNullOrEmpty(formatString))
			{
				return null;
			}

			return FormatCellUsingFormatString(cell, formatString);
		}
		catch
		{
			// Results in a string value
			return null;
		}
	}
}
