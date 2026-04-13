namespace PanoramicData.SheetMagic;

/// <summary>
/// Defines a single conditional formatting rule with its condition and style.
/// Maps to an Excel ConditionalFormattingRule element.
/// </summary>
/// <example>
/// <code>
/// var rule = new ConditionalFormatRule
/// {
///     RuleType = ConditionalFormatRuleType.CellIs,
///     Operator = ConditionalFormatOperator.GreaterThan,
///     Formula = "5",
///     Style = new ConditionalFormatStyle
///     {
///         FontColor = System.Drawing.Color.Green,
///         FontWeight = FontWeight.Bold
///     }
/// };
/// </code>
/// </example>
public class ConditionalFormatRule
{
	/// <summary>
	/// The type of conditional formatting rule.
	/// </summary>
	public ConditionalFormatRuleType RuleType { get; set; }

	/// <summary>
	/// Operator for CellIs rules. Required when RuleType is CellIs.
	/// </summary>
	public ConditionalFormatOperator? Operator { get; set; }

	/// <summary>
	/// Formula or comparison value. For CellIs rules, this is the value to compare against.
	/// For Expression rules, this is the full formula. For Between/NotBetween, this is the lower bound.
	/// </summary>
	/// <remarks>
	/// For expression rules, write the formula exactly as Excel would store it for the top-left data cell in the target range,
	/// for example <c>LEN(A2)=0</c> or <c>AND(C2&gt;=1,C2&lt;=10)</c>.
	/// </remarks>
	public string? Formula { get; set; }

	/// <summary>
	/// Second formula for Between/NotBetween operators (the upper bound).
	/// </summary>
	public string? Formula2 { get; set; }

	/// <summary>
	/// Text value for text-based rules (ContainsText, NotContainsText, BeginsWith, EndsWith).
	/// </summary>
	public string? Text { get; set; }

	/// <summary>
	/// Rank value for Top10 rules. Defaults to 10.
	/// </summary>
	public uint? Rank { get; set; }

	/// <summary>
	/// If true, Top10 rule selects bottom values instead of top. Defaults to false.
	/// </summary>
	public bool Bottom { get; set; }

	/// <summary>
	/// If true, Top10 rank is a percentage rather than an absolute count. Defaults to false.
	/// </summary>
	public bool Percent { get; set; }

	/// <summary>
	/// For AboveAverage rules: true = above average, false = below average. Defaults to true.
	/// </summary>
	public bool AboveAverage { get; set; } = true;

	/// <summary>
	/// For AboveAverage rules: if true, includes values equal to the average. Defaults to false.
	/// </summary>
	public bool EqualAverage { get; set; }

	/// <summary>
	/// If true, no subsequent rules are evaluated when this rule matches. Defaults to false.
	/// </summary>
	public bool StopIfTrue { get; set; }

	/// <summary>
	/// The formatting style to apply when this rule's condition is met.
	/// </summary>
	/// <example>
	/// <code>
	/// Style = new ConditionalFormatStyle
	/// {
	///     BackgroundColor = System.Drawing.Color.Red
	/// };
	/// </code>
	/// </example>
	public ConditionalFormatStyle Style { get; set; } = new();

	internal void Validate()
	{
		if (!Style.HasFormatting())
		{
			throw new ValidationException($"{nameof(ConditionalFormatRule)} must define at least one style property.");
		}

		switch (RuleType)
		{
			case ConditionalFormatRuleType.CellIs:
				if (Operator is null)
				{
					throw new ValidationException($"{nameof(ConditionalFormatRuleType.CellIs)} rules require {nameof(Operator)}.");
				}

				if (string.IsNullOrWhiteSpace(Formula))
				{
					throw new ValidationException($"{nameof(ConditionalFormatRuleType.CellIs)} rules require {nameof(Formula)}.");
				}

				if (Operator is ConditionalFormatOperator.Between or ConditionalFormatOperator.NotBetween && string.IsNullOrWhiteSpace(Formula2))
				{
					throw new ValidationException($"{Operator} rules require {nameof(Formula2)}.");
				}

				break;

			case ConditionalFormatRuleType.Expression:
				if (string.IsNullOrWhiteSpace(Formula))
				{
					throw new ValidationException($"{nameof(ConditionalFormatRuleType.Expression)} rules require {nameof(Formula)}.");
				}

				break;

			case ConditionalFormatRuleType.ContainsText:
			case ConditionalFormatRuleType.NotContainsText:
			case ConditionalFormatRuleType.BeginsWith:
			case ConditionalFormatRuleType.EndsWith:
				if (string.IsNullOrWhiteSpace(Text))
				{
					throw new ValidationException($"{RuleType} rules require {nameof(Text)}.");
				}

				break;

			case ConditionalFormatRuleType.Top10:
				if (Rank == 0)
				{
					throw new ValidationException($"{nameof(ConditionalFormatRuleType.Top10)} rules require {nameof(Rank)} to be greater than zero.");
				}

				break;
		}
	}
}
