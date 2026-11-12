namespace PanoramicData.SheetMagic;

/// <summary>
/// Sheet name determination and validation
/// </summary>
public partial class MagicSpreadsheet
{
	private const int SheetNameCharacterLimit = 31;

	private string DetermineAndValidateSheetName(string? sheetName, string typeName)
	{
		sheetName = DetermineSheetName(sheetName, typeName);
		ValidateSheetNameLength(sheetName);
		ValidateSheetNameUniqueness(sheetName);
		return sheetName;
	}

	private static string DetermineSheetName(string? sheetName, string typeName)
	{
		if (!string.IsNullOrWhiteSpace(sheetName))
		{
			return sheetName!;
		}

		try
		{
			// Get the length and leave space for an "s" on the end
			var length = Math.Min(typeName.Length, SheetNameCharacterLimit - 1);
			return $"{typeName[..length]}s";
		}
		catch (Exception)
		{
			return "Sheet";
		}
	}

	private static void ValidateSheetNameLength(string sheetName)
	{
		if (sheetName.Length > SheetNameCharacterLimit)
		{
			throw new ArgumentException($"Sheet name cannot be more than {SheetNameCharacterLimit} characters", nameof(sheetName));
		}
	}

	private void ValidateSheetNameUniqueness(string sheetName)
	{
		var sheetExists = _document!.WorkbookPart!.Workbook!.Sheets?
			.Any(existingSheet => string.Equals(((Sheet)existingSheet).Name!.Value, sheetName, StringComparison.InvariantCultureIgnoreCase)) ?? false;

		if (sheetExists)
		{
			throw new ArgumentException($"Sheet name {sheetName} already exists. Sheet names must be unique per Workbook", nameof(sheetName));
		}
	}
}
