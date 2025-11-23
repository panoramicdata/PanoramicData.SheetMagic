using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;

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
			return actualDate.Value.ToString(
				formatString
					.Replace("\\", string.Empty)
					.Replace(";", string.Empty)
					.Replace("@", string.Empty)
					.Replace("m", "M"))
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

	private static bool HasStyleIndex(Cell cell)
		=> cell.StyleIndex?.HasValue == true;

	private (CellFormats? cellFormats, NumberingFormats? numberingFormats) GetFormattingParts()
	{
		var cellFormats = _document?.WorkbookPart?.WorkbookStylesPart?.Stylesheet.CellFormats;
		var numberingFormats = _document?.WorkbookPart?.WorkbookStylesPart?.Stylesheet.NumberingFormats;
		return (cellFormats, numberingFormats);
	}

	private static string? GetFormatString(uint numberFormatId, NumberingFormats? numberingFormats)
	{
		if (numberFormatId >= BuiltInCellFormats.CustomFormatStartIndex)
		{
			return GetCustomFormatString(numberFormatId, numberingFormats);
		}
		else
		{
			return GetBuiltInFormatString((int)numberFormatId);
		}
	}

	private static string? GetCustomFormatString(uint numberFormatId, NumberingFormats? numberingFormats)
	{
		if (numberingFormats == null)
		{
			return null;
		}

		var numberingFormat = numberingFormats
			.Cast<NumberingFormat>()
			.SingleOrDefault(f => f.NumberFormatId?.Value == numberFormatId);

		return numberingFormat?.FormatCode?.Value;
	}

	private static string? GetBuiltInFormatString(int numberFormatId)
	{
		var builtInFormat = BuiltInCellFormats.GetBuiltInCellFormatByIndex(numberFormatId);

		if (builtInFormat == null)
		{
			return null;
		}

		if (builtInFormat.Value.formatType is
			CellFormatType.General or
			CellFormatType.Text or
			CellFormatType.Unknown)
		{
			// Results in a string
			return null;
		}

		return builtInFormat.Value.formatString;
	}

	private static string? FormatCellUsingFormatString(Cell cell, string formatString)
		=> IsFormatStringADate(formatString)
			? FormatCellAsDateTime(cell, formatString)
			: FormatCellAsNumber(cell, formatString);
}
