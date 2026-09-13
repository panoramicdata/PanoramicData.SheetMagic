namespace PanoramicData.SheetMagic;

/// <summary>
/// Options for configuring how a sheet is added to a spreadsheet.
/// </summary>
/// <example>
/// <code>
/// var options = new AddSheetOptions
/// {
///     ConditionalFormats =
///     [
///         new ConditionalFormat
///         {
///             ColumnNames = ["Score"],
///             Rules =
///             [
///                 new ConditionalFormatRule
///                 {
///                     RuleType = ConditionalFormatRuleType.ContainsBlanks,
///                     Style = new ConditionalFormatStyle
///                     {
///                         BackgroundColor = System.Drawing.Color.Red
///                     }
///                 },
///                 new ConditionalFormatRule
///                 {
///                     RuleType = ConditionalFormatRuleType.CellIs,
///                     Operator = ConditionalFormatOperator.GreaterThan,
///                     Formula = "5",
///                     Style = new ConditionalFormatStyle
///                     {
///                         FontColor = System.Drawing.Color.Green,
///                         FontWeight = FontWeight.Bold
///                     }
///                 }
///             ]
///         }
///     ]
/// };
/// </code>
/// </example>
public class AddSheetOptions
{
	/// <summary>
	/// The properties to include
	/// </summary>
	public HashSet<string>? IncludeProperties { get; set; }

	/// <summary>
	/// The properties to exclude
	/// </summary>
	public HashSet<string>? ExcludeProperties { get; set; }

	/// <summary>
	/// The order properties should be output.
	/// </summary>
	public string[]? PropertyOrder { get; set; }

	/// <summary>
	/// Explicit header text for properties.
	/// </summary>
	public string[]? PropertyHeaders { get; set; }

	/// <summary>
	/// Whether to sort the combined list of properties, and any additional extended properties. Defaults to true.
	/// </summary>
	public bool SortExtendedProperties { get; set; } = true;

	/// <summary>
	/// TableOptions
	/// </summary>
	public TableOptions? TableOptions { get; set; } = new TableOptions
	{
		XlsxTableStyle = XlsxTableStyle.TableStyleMedium11
	};

	/// <summary>
	/// An optional EnumerableCellOptions.  If not set, the Options EnumerableCellOptions set in Options is used.
	/// </summary>
	public EnumerableCellOptions? EnumerableCellOptions { get; set; }

	/// <summary>
	/// In Excel, it is not possible to add a table with no rows.
	/// If the user tries to add a table with no rows and this property is set to:
	/// - true (default): SheetMagic will throw an InvalidOperationException if
	/// - false: SheetMagic will silently not add a new sheet
	/// </summary>
	public bool ThrowExceptionOnEmptyList { get; set; } = true;

	/// <summary>
	/// Optional list of conditional formatting specifications to apply to the sheet.
	/// Each ConditionalFormat can target specific columns and contain multiple rules.
	/// </summary>
	/// <remarks>
	/// Column names must match the final header text written to Excel.
	/// If <see cref="PropertyHeaders"/> is set, use those values.
	/// Otherwise use the property's <c>Description</c> attribute value, or the property name when no description is present.
	/// Leave <see cref="ConditionalFormat.ColumnNames"/> empty to apply a conditional format to every exported column.
	/// </remarks>
	/// <example>
	/// <code>
	/// var options = new AddSheetOptions
	/// {
	///     PropertyHeaders = ["Name", "Description", "Score"],
	///     ConditionalFormats =
	///     [
	///         new ConditionalFormat
	///         {
	///             ColumnNames = ["Name", "Description"],
	///             Rules =
	///             [
	///                 new ConditionalFormatRule
	///                 {
	///                     RuleType = ConditionalFormatRuleType.ContainsBlanks,
	///                     Style = new ConditionalFormatStyle
	///                     {
	///                         BackgroundColor = System.Drawing.Color.Red
	///                     }
	///                 }
	///             ]
	///         }
	///     ]
	/// };
	/// </code>
	/// </example>
	public List<ConditionalFormat>? ConditionalFormats { get; set; }

	/// <summary>
	/// Validates the options configuration.
	/// </summary>
	/// <param name="tableStyles">The list of custom table styles to validate against.</param>
	/// <exception cref="ValidationException">Thrown when validation fails.</exception>
	public void Validate(List<CustomTableStyle> tableStyles)
	{
		if (IncludeProperties != null && ExcludeProperties != null)
		{
			throw new ValidationException($"Cannot set both {nameof(IncludeProperties)} and {nameof(ExcludeProperties)}");
		}

		if (ConditionalFormats is not null)
		{
			foreach (var conditionalFormat in ConditionalFormats)
			{
				conditionalFormat.Validate();
			}
		}

		TableOptions?.Validate(tableStyles);
	}

	internal AddSheetOptions Clone()
		=> new()
		{
			EnumerableCellOptions = CloneEnumerableCellOptions(EnumerableCellOptions),
			ExcludeProperties = ExcludeProperties == null
				? null
				: [.. ExcludeProperties],
			IncludeProperties = IncludeProperties == null
				? null
				: [.. IncludeProperties],
			PropertyOrder = PropertyOrder,
			PropertyHeaders = PropertyHeaders,
			SortExtendedProperties = SortExtendedProperties,
			TableOptions = CloneTableOptions(TableOptions),
			ThrowExceptionOnEmptyList = ThrowExceptionOnEmptyList,
			ConditionalFormats = ConditionalFormats?.Select(CloneConditionalFormat).ToList()
		};

	private static EnumerableCellOptions? CloneEnumerableCellOptions(EnumerableCellOptions? enumerableCellOptions)
		=> enumerableCellOptions is null
			? null
			: new EnumerableCellOptions
			{
				CellDelimiter = enumerableCellOptions.CellDelimiter,
				Expand = enumerableCellOptions.Expand,
			};

	private static TableOptions? CloneTableOptions(TableOptions? tableOptions)
		=> tableOptions is null
			? null
			: new TableOptions
			{
				CustomTableStyle = tableOptions.CustomTableStyle,
				DisplayName = tableOptions.DisplayName,
				Name = tableOptions.Name,
				ShowColumnStripes = tableOptions.ShowColumnStripes,
				ShowFirstColumn = tableOptions.ShowFirstColumn,
				ShowLastColumn = tableOptions.ShowLastColumn,
				ShowRowStripes = tableOptions.ShowRowStripes,
				ShowTotalsRow = tableOptions.ShowTotalsRow,
				XlsxTableStyle = tableOptions.XlsxTableStyle
			};

	private static ConditionalFormat CloneConditionalFormat(ConditionalFormat conditionalFormat)
		=> new()
		{
			ColumnNames = conditionalFormat.ColumnNames is null ? null : [.. conditionalFormat.ColumnNames],
			Rules = [.. conditionalFormat.Rules.Select(CloneConditionalFormatRule)]
		};

	private static ConditionalFormatRule CloneConditionalFormatRule(ConditionalFormatRule rule)
		=> new()
		{
			RuleType = rule.RuleType,
			Operator = rule.Operator,
			Formula = rule.Formula,
			Formula2 = rule.Formula2,
			Text = rule.Text,
			Rank = rule.Rank,
			Bottom = rule.Bottom,
			Percent = rule.Percent,
			AboveAverage = rule.AboveAverage,
			EqualAverage = rule.EqualAverage,
			StopIfTrue = rule.StopIfTrue,
			Style = CloneConditionalFormatStyle(rule.Style)
		};

	private static ConditionalFormatStyle CloneConditionalFormatStyle(ConditionalFormatStyle style)
		=> new()
		{
			FontColor = style.FontColor,
			FontWeight = style.FontWeight,
			Italic = style.Italic,
			Strikethrough = style.Strikethrough,
			BackgroundColor = style.BackgroundColor,
			BorderColor = style.BorderColor,
			NumberFormat = style.NumberFormat
		};
}