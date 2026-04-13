namespace PanoramicData.SheetMagic;

/// <summary>
/// Operator for CellIs conditional formatting rules, mapping to Excel's conditional formatting operators.
/// </summary>
public enum ConditionalFormatOperator
{
	/// <summary>Cell value equals the formula value.</summary>
	Equal,

	/// <summary>Cell value does not equal the formula value.</summary>
	NotEqual,

	/// <summary>Cell value is greater than the formula value.</summary>
	GreaterThan,

	/// <summary>Cell value is greater than or equal to the formula value.</summary>
	GreaterThanOrEqual,

	/// <summary>Cell value is less than the formula value.</summary>
	LessThan,

	/// <summary>Cell value is less than or equal to the formula value.</summary>
	LessThanOrEqual,

	/// <summary>Cell value is between Formula and Formula2 (inclusive).</summary>
	Between,

	/// <summary>Cell value is not between Formula and Formula2.</summary>
	NotBetween
}
