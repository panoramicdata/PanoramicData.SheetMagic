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
			EnumerableCellOptions = EnumerableCellOptions == null
				? null
				: new EnumerableCellOptions
				{
					CellDelimiter = EnumerableCellOptions.CellDelimiter,
					Expand = EnumerableCellOptions.Expand,
				},
			ExcludeProperties = ExcludeProperties == null
				? null
				: [.. ExcludeProperties],
			IncludeProperties = IncludeProperties == null
				? null
				: [.. IncludeProperties],
			PropertyOrder = PropertyOrder,
			PropertyHeaders = PropertyHeaders,
			SortExtendedProperties = SortExtendedProperties,
			TableOptions = TableOptions == null
				? null
				: new TableOptions
				{
					CustomTableStyle = TableOptions.CustomTableStyle,
					DisplayName = TableOptions.DisplayName,
					Name = TableOptions.Name,
					ShowColumnStripes = TableOptions.ShowColumnStripes,
					ShowFirstColumn = TableOptions.ShowFirstColumn,
					ShowLastColumn = TableOptions.ShowLastColumn,
					ShowRowStripes = TableOptions.ShowRowStripes,
					ShowTotalsRow = TableOptions.ShowTotalsRow,
					XlsxTableStyle = TableOptions.XlsxTableStyle
				},
			ThrowExceptionOnEmptyList = ThrowExceptionOnEmptyList,
			ConditionalFormats = ConditionalFormats?.Select(cf => new ConditionalFormat
			{
				ColumnNames = cf.ColumnNames is null ? null : [.. cf.ColumnNames],
				Rules = [.. cf.Rules.Select(r => new ConditionalFormatRule
				{
					RuleType = r.RuleType,
					Operator = r.Operator,
					Formula = r.Formula,
					Formula2 = r.Formula2,
					Text = r.Text,
					Rank = r.Rank,
					Bottom = r.Bottom,
					Percent = r.Percent,
					AboveAverage = r.AboveAverage,
					EqualAverage = r.EqualAverage,
					StopIfTrue = r.StopIfTrue,
					Style = new ConditionalFormatStyle
					{
						FontColor = r.Style.FontColor,
						FontWeight = r.Style.FontWeight,
						Italic = r.Style.Italic,
						Strikethrough = r.Style.Strikethrough,
						BackgroundColor = r.Style.BackgroundColor,
						BorderColor = r.Style.BorderColor,
						NumberFormat = r.Style.NumberFormat
					}
				})]
			}).ToList()
		};
}