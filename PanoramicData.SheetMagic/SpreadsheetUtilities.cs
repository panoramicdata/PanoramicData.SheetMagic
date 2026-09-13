using System.Text;

namespace PanoramicData.SheetMagic;

/// <summary>
/// Utility methods for translating between Excel cell references and indices, and for
/// normalising text when matching sheet, column and property names.
/// </summary>
internal static class SpreadsheetUtilities
{
	private const string Letters = "abcdefghijklmnopqrstuvwxyz";
	private const string Numbers = "0123456789";

	// A plain Regex rather than a [GeneratedRegex] partial method: this type is deliberately not
	// partial, so that it reads as the single, self-contained helper class that it is.
	private static readonly Regex CellReferenceRegex = new(@"(?<col>([A-Z]|[a-z])+)(?<row>(\d)+)", RegexOptions.Compiled);

	internal static string ColumnLetter(int intCol)
	{
		var intFirstLetter = (intCol / 676) + 64;
		var intSecondLetter = (intCol % 676 / 26) + 64;
		var intThirdLetter = (intCol % 26) + 65;

		var firstLetter = intFirstLetter > 64
			 ? (char)intFirstLetter
			 : ' ';
		var secondLetter = intSecondLetter > 64
			 ? (char)intSecondLetter
			 : ' ';
		var thirdLetter = (char)intThirdLetter;

		return string.Concat(firstLetter, secondLetter,
			 thirdLetter).Trim();
	}

	internal static bool StringsMatch(string string1, string string2) => TweakString(string1) == TweakString(string2);

	internal static string TweakString(string text)
	{
		var stringBuilder = new StringBuilder();

		foreach (var @char in text.ToLowerInvariant())
		{
			if (!Letters.Contains(@char) && !Numbers.Contains(@char))
			{
				continue;
			}

			_ = stringBuilder.Append(@char);
		}

		var tweakString = stringBuilder.ToString();

		// Chop numbers from the beginning
		while (tweakString.Length > 0 && Numbers.Contains(tweakString[0]))
		{
			tweakString = tweakString[1..];
		}

		// Remove plurals
		return tweakString.EndsWith('s') && !tweakString.EndsWith("ss")
			 ? tweakString[..^1]
			 : tweakString;
	}

	internal static (int columnIndex, int rowIndex) GetReference(string cellReference)
	{
		var match = CellReferenceRegex.Match(cellReference)
			?? throw new ArgumentException($"Invalid cell reference {cellReference}", nameof(cellReference));

		var col = match.Groups["col"].Value;
		var row = match.Groups["row"].Value;

		return (ExcelColumnNameToNumber(col) - 1, int.Parse(row) - 1);
	}

	private static int ExcelColumnNameToNumber(string columnName)
	{
		if (string.IsNullOrEmpty(columnName))
		{
			throw new ArgumentNullException(nameof(columnName));
		}

		var sum = 0;
		foreach (var t in columnName.ToUpperInvariant())
		{
			sum *= 26;
			sum += t - 'A' + 1;
		}

		return sum;
	}
}
