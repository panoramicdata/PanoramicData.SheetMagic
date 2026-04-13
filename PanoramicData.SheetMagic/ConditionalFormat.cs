namespace PanoramicData.SheetMagic;

/// <summary>
/// Defines conditional formatting to apply to specific columns (or all columns) in a sheet.
/// Contains one or more rules that are evaluated in priority order.
/// </summary>
/// <example>
/// <code>
/// var format = new ConditionalFormat
/// {
///     ColumnNames = ["Score"],
///     Rules =
///     [
///         new ConditionalFormatRule
///         {
///             RuleType = ConditionalFormatRuleType.CellIs,
///             Operator = ConditionalFormatOperator.GreaterThan,
///             Formula = "5",
///             Style = new ConditionalFormatStyle
///             {
///                 FontColor = System.Drawing.Color.Green,
///                 FontWeight = FontWeight.Bold
///             }
///         }
///     ]
/// };
/// </code>
/// </example>
public class ConditionalFormat
{
	/// <summary>
	/// The column header names to apply the formatting to.
	/// If null or empty, the formatting applies to all data columns.
	/// Column names should match the final header text shown in Excel.
	/// This means <see cref="AddSheetOptions.PropertyHeaders"/> values when present,
	/// otherwise the property's <c>Description</c> attribute value, or the property name if no description is set.
	/// </summary>
	public List<string>? ColumnNames { get; set; }

	/// <summary>
	/// The conditional formatting rules to evaluate, in priority order.
	/// </summary>
	/// <remarks>
	/// Each rule maps closely to Excel's conditional formatting model and becomes a separate OpenXML conditional formatting rule.
	/// </remarks>
	public List<ConditionalFormatRule> Rules { get; set; } = [];

	internal void Validate()
	{
		if (Rules.Count == 0)
		{
			throw new ValidationException($"{nameof(ConditionalFormat)} must contain at least one rule.");
		}

		foreach (var rule in Rules)
		{
			rule.Validate();
		}
	}
}
