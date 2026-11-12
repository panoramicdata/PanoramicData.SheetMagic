namespace PanoramicData.SheetMagic;

/// <summary>
/// Conditional formatting rule construction and enum mapping
/// </summary>
public partial class MagicSpreadsheet
{
	private static ConditionalFormattingRule BuildConditionalFormattingRule(
		ConditionalFormatRule rule,
		uint dxfId,
		int priority,
		string firstCellRef)
	{
		rule.Validate();

		var cfRule = new ConditionalFormattingRule
		{
			Type = MapConditionalFormatRuleType(rule.RuleType),
			FormatId = dxfId,
			Priority = priority
		};

		if (rule.StopIfTrue)
		{
			cfRule.StopIfTrue = true;
		}

		ConfigureConditionalFormattingRule(cfRule, rule, firstCellRef);

		return cfRule;
	}

	/// <summary>
	/// Applies the parts of the rule that differ by rule type: the comparison operator, the text
	/// being matched and the formula Excel evaluates for each cell.
	/// </summary>
	private static void ConfigureConditionalFormattingRule(
		ConditionalFormattingRule cfRule,
		ConditionalFormatRule rule,
		string firstCellRef)
	{
		if (TryConfigureCellIsRule(cfRule, rule))
		{
			return;
		}

		if (TryConfigureFormulaRule(cfRule, rule, firstCellRef))
		{
			return;
		}

		if (TryConfigureTextRule(cfRule, rule, firstCellRef))
		{
			return;
		}

		ConfigureRankingRule(cfRule, rule);
	}

	private static bool TryConfigureCellIsRule(ConditionalFormattingRule cfRule, ConditionalFormatRule rule)
	{
		if (rule.RuleType != ConditionalFormatRuleType.CellIs)
		{
			return false;
		}

		cfRule.Operator = MapConditionalFormatOperator(rule.Operator ?? throw new ValidationException($"{nameof(ConditionalFormatRuleType.CellIs)} rules require {nameof(rule.Operator)}."));
		cfRule.Append(new Formula(rule.Formula!));

		if (rule.Formula2 is not null)
		{
			cfRule.Append(new Formula(rule.Formula2));
		}

		return true;
	}

	/// <summary>
	/// Configures the rule types whose whole configuration is a single formula.
	/// </summary>
	private static bool TryConfigureFormulaRule(
		ConditionalFormattingRule cfRule,
		ConditionalFormatRule rule,
		string firstCellRef)
	{
		var formula = rule.RuleType switch
		{
			ConditionalFormatRuleType.Expression => rule.Formula!,
			ConditionalFormatRuleType.ContainsBlanks => $"LEN(TRIM({firstCellRef}))=0",
			ConditionalFormatRuleType.NotContainsBlanks => $"LEN(TRIM({firstCellRef}))>0",
			ConditionalFormatRuleType.ContainsErrors => $"ISERROR({firstCellRef})",
			ConditionalFormatRuleType.NotContainsErrors => $"NOT(ISERROR({firstCellRef}))",
			_ => null
		};

		if (formula is null)
		{
			return false;
		}

		cfRule.Append(new Formula(formula));
		return true;
	}

	/// <summary>
	/// Configures the rule types that match against <see cref="ConditionalFormatRule.Text"/>.
	/// </summary>
	private static bool TryConfigureTextRule(
		ConditionalFormattingRule cfRule,
		ConditionalFormatRule rule,
		string firstCellRef)
	{
		var text = rule.Text;
		var escapedText = text is null ? null : EscapeFormulaText(text);

		(ConditionalFormattingOperatorValues Operator, string? Formula) configuration = rule.RuleType switch
		{
			ConditionalFormatRuleType.ContainsText =>
				(ConditionalFormattingOperatorValues.ContainsText, $"NOT(ISERROR(SEARCH(\"{escapedText}\",{firstCellRef})))"),
			ConditionalFormatRuleType.NotContainsText =>
				(ConditionalFormattingOperatorValues.NotContains, $"ISERROR(SEARCH(\"{escapedText}\",{firstCellRef}))"),
			ConditionalFormatRuleType.BeginsWith =>
				(ConditionalFormattingOperatorValues.BeginsWith, $"LEFT({firstCellRef},{text!.Length})=\"{escapedText}\""),
			ConditionalFormatRuleType.EndsWith =>
				(ConditionalFormattingOperatorValues.EndsWith, $"RIGHT({firstCellRef},{text!.Length})=\"{escapedText}\""),
			_ => (default, null)
		};

		if (configuration.Formula is null)
		{
			return false;
		}

		cfRule.Operator = configuration.Operator;
		cfRule.Text = text;
		cfRule.Append(new Formula(configuration.Formula));
		return true;
	}

	/// <summary>
	/// Configures the rule types that rank values against the other values in the column.
	/// </summary>
	private static void ConfigureRankingRule(ConditionalFormattingRule cfRule, ConditionalFormatRule rule)
	{
		switch (rule.RuleType)
		{
			case ConditionalFormatRuleType.Top10:
				cfRule.Rank = rule.Rank ?? 10;
				cfRule.Bottom = rule.Bottom;
				cfRule.Percent = rule.Percent;
				break;

			case ConditionalFormatRuleType.AboveAverage:
				if (!rule.AboveAverage)
				{
					cfRule.AboveAverage = false;
				}

				if (rule.EqualAverage)
				{
					cfRule.EqualAverage = true;
				}

				break;

			default:
				// DuplicateValues and UniqueValues need no further configuration
				break;
		}
	}

	private static ConditionalFormatValues MapConditionalFormatRuleType(ConditionalFormatRuleType ruleType)
		=> ruleType switch
		{
			ConditionalFormatRuleType.CellIs => ConditionalFormatValues.CellIs,
			ConditionalFormatRuleType.Expression => ConditionalFormatValues.Expression,
			ConditionalFormatRuleType.ContainsBlanks => ConditionalFormatValues.ContainsBlanks,
			ConditionalFormatRuleType.NotContainsBlanks => ConditionalFormatValues.NotContainsBlanks,
			ConditionalFormatRuleType.ContainsErrors => ConditionalFormatValues.ContainsErrors,
			ConditionalFormatRuleType.NotContainsErrors => ConditionalFormatValues.NotContainsErrors,
			ConditionalFormatRuleType.ContainsText => ConditionalFormatValues.ContainsText,
			ConditionalFormatRuleType.NotContainsText => ConditionalFormatValues.NotContainsText,
			ConditionalFormatRuleType.BeginsWith => ConditionalFormatValues.BeginsWith,
			ConditionalFormatRuleType.EndsWith => ConditionalFormatValues.EndsWith,
			ConditionalFormatRuleType.DuplicateValues => ConditionalFormatValues.DuplicateValues,
			ConditionalFormatRuleType.UniqueValues => ConditionalFormatValues.UniqueValues,
			ConditionalFormatRuleType.Top10 => ConditionalFormatValues.Top10,
			ConditionalFormatRuleType.AboveAverage => ConditionalFormatValues.AboveAverage,
			_ => throw new ArgumentOutOfRangeException(nameof(ruleType), ruleType, "Unsupported conditional format rule type.")
		};

	private static ConditionalFormattingOperatorValues MapConditionalFormatOperator(ConditionalFormatOperator op)
		=> op switch
		{
			ConditionalFormatOperator.Equal => ConditionalFormattingOperatorValues.Equal,
			ConditionalFormatOperator.NotEqual => ConditionalFormattingOperatorValues.NotEqual,
			ConditionalFormatOperator.GreaterThan => ConditionalFormattingOperatorValues.GreaterThan,
			ConditionalFormatOperator.GreaterThanOrEqual => ConditionalFormattingOperatorValues.GreaterThanOrEqual,
			ConditionalFormatOperator.LessThan => ConditionalFormattingOperatorValues.LessThan,
			ConditionalFormatOperator.LessThanOrEqual => ConditionalFormattingOperatorValues.LessThanOrEqual,
			ConditionalFormatOperator.Between => ConditionalFormattingOperatorValues.Between,
			ConditionalFormatOperator.NotBetween => ConditionalFormattingOperatorValues.NotBetween,
			_ => throw new ArgumentOutOfRangeException(nameof(op), op, "Unsupported conditional format operator.")
		};
}
