namespace PanoramicData.SheetMagic;

/// <summary>
/// The type of conditional formatting rule, mapping to Excel's conditional format types.
/// </summary>
public enum ConditionalFormatRuleType
{
	/// <summary>Cell value comparison (uses Operator and Formula/Formula2).</summary>
	CellIs,

	/// <summary>Custom formula expression.</summary>
	Expression,

	/// <summary>Cell contains blank/empty value.</summary>
	ContainsBlanks,

	/// <summary>Cell does not contain blank/empty value.</summary>
	NotContainsBlanks,

	/// <summary>Cell contains an error.</summary>
	ContainsErrors,

	/// <summary>Cell does not contain an error.</summary>
	NotContainsErrors,

	/// <summary>Cell text contains specified text (uses Text property).</summary>
	ContainsText,

	/// <summary>Cell text does not contain specified text (uses Text property).</summary>
	NotContainsText,

	/// <summary>Cell text begins with specified text (uses Text property).</summary>
	BeginsWith,

	/// <summary>Cell text ends with specified text (uses Text property).</summary>
	EndsWith,

	/// <summary>Duplicate values in the range.</summary>
	DuplicateValues,

	/// <summary>Unique values in the range.</summary>
	UniqueValues,

	/// <summary>Top or bottom N values (uses Rank, Bottom, Percent).</summary>
	Top10,

	/// <summary>Values above or below average (uses AboveAverage, EqualAverage).</summary>
	AboveAverage
}
