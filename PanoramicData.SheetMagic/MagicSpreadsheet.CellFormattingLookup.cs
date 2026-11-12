namespace PanoramicData.SheetMagic;

/// <summary>
/// Number format lookup for cell styles
/// </summary>
public partial class MagicSpreadsheet
{
	private static bool HasStyleIndex(Cell cell)
		=> cell.StyleIndex?.HasValue == true;

	private (CellFormats? cellFormats, NumberingFormats? numberingFormats) GetFormattingParts()
	{
		var stylesheet = GetStylesheet();
		return (stylesheet?.CellFormats, stylesheet?.NumberingFormats);
	}

	private Stylesheet? GetStylesheet()
		=> _document?.WorkbookPart?.WorkbookStylesPart?.Stylesheet;

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
