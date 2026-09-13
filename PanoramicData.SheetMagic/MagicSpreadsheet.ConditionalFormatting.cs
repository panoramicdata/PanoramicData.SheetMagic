using DocumentFormat.OpenXml;
using PanoramicData.SheetMagic.Extensions;

namespace PanoramicData.SheetMagic;

/// <summary>
/// Conditional formatting methods for applying Excel conditional formats to worksheets.
/// </summary>
public partial class MagicSpreadsheet
{
	internal void ApplyConditionalFormatting(
		AddSheetOptions addSheetOptions,
		WorksheetPart worksheetPart,
		List<PropertyInfo> propertyList,
		List<string> keyList,
		int rowCount)
	{
		if (addSheetOptions.ConditionalFormats is null || addSheetOptions.ConditionalFormats.Count == 0)
		{
			return;
		}

		var columnHeaders = propertyList
			.Select(p => p.GetPropertyDescription() ?? p.Name)
			.Concat(keyList)
			.ToList();

		var stylesheet = _document!.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
		var differentialFormats = stylesheet.Descendants<DifferentialFormats>().FirstOrDefault();
		if (differentialFormats is null)
		{
			differentialFormats = new DifferentialFormats();
			stylesheet.Append(differentialFormats);
		}

		var currentDxfId = (uint)differentialFormats.Elements<DifferentialFormat>().Count();
		var nextNumberFormatId = GetNextConditionalNumberFormatId(stylesheet);
		var priority = 1;

		foreach (var conditionalFormat in addSheetOptions.ConditionalFormats)
		{
			var targetColumnIndices = ResolveTargetColumnIndices(conditionalFormat.ColumnNames, columnHeaders);

			foreach (var targetColumnIndex in targetColumnIndices.Distinct())
			{
				var sqRef = BuildConditionalFormattingSqRef(targetColumnIndex, rowCount);
				var firstCellRef = BuildFirstCellReference(targetColumnIndex);
				var cfElement = new ConditionalFormatting
				{
					SequenceOfReferences = new ListValue<StringValue>([new StringValue(sqRef)])
				};

				foreach (var rule in conditionalFormat.Rules)
				{
					var dxf = CreateDifferentialFormatFromStyle(rule.Style, ref nextNumberFormatId);
					differentialFormats.Append(dxf);

					var cfRule = BuildConditionalFormattingRule(rule, currentDxfId, priority++, firstCellRef);
					cfElement.Append(cfRule);
					currentDxfId++;
				}

				worksheetPart.Worksheet!.Append(cfElement);
			}
		}

		differentialFormats.Count = currentDxfId;
	}

	private static List<int> ResolveTargetColumnIndices(List<string>? columnNames, List<string> allColumnHeaders)
	{
		if (columnNames is null || columnNames.Count == 0)
		{
			return Enumerable.Range(0, allColumnHeaders.Count).ToList();
		}

		var indices = new List<int>();
		foreach (var name in columnNames)
		{
			var index = allColumnHeaders.IndexOf(name);
			if (index < 0)
			{
				throw new ArgumentException(
					$"Conditional format column '{name}' not found in sheet headers. " +
					$"Available columns: {string.Join(", ", allColumnHeaders)}");
			}

			indices.Add(index);
		}

		return indices;
	}

	private static string BuildConditionalFormattingSqRef(int columnIndex, int rowCount)
	{
		// Data starts at row 2 (row 1 is the header row)
		var lastRow = rowCount + 1;
		var colLetter = ColumnLetter(columnIndex);
		return $"{colLetter}2:{colLetter}{lastRow}";
	}

	private static string BuildFirstCellReference(int columnIndex)
	{
		var firstColLetter = ColumnLetter(columnIndex);
		return $"{firstColLetter}2";
	}

	private static DifferentialFormat CreateDifferentialFormatFromStyle(ConditionalFormatStyle style, ref uint nextNumberFormatId)
	{
		var dxf = new DifferentialFormat();

		var font = CreateDifferentialFont(style);
		if (font is not null)
		{
			dxf.Append(font);
		}

		if (style.BackgroundColor.HasValue)
		{
			dxf.Append(CreateSolidFill(style.BackgroundColor.Value));
		}

		if (style.BorderColor.HasValue)
		{
			dxf.Append(CreateThinBorder(style.BorderColor.Value));
		}

		if (style.NumberFormat is not null)
		{
			dxf.Append(new NumberingFormat
			{
				NumberFormatId = nextNumberFormatId++,
				FormatCode = style.NumberFormat
			});
		}

		return dxf;
	}

	private static Font? CreateDifferentialFont(ConditionalFormatStyle style)
	{
		if (!HasFontStyling(style))
		{
			return null;
		}

		var font = new Font();
		if (style.FontWeight == FontWeight.Bold)
		{
			font.Append(new Bold());
		}

		if (style.Italic == true)
		{
			font.Append(new Italic());
		}

		if (style.Strikethrough == true)
		{
			font.Append(new Strike());
		}

		if (style.FontColor.HasValue)
		{
			font.Append(GetColor(style.FontColor.Value));
		}

		return font;
	}

	private static bool HasFontStyling(ConditionalFormatStyle style)
		=> style.FontColor.HasValue
			|| style.FontWeight.HasValue
			|| style.Italic.HasValue
			|| style.Strikethrough.HasValue;

	private static Fill CreateSolidFill(System.Drawing.Color color)
	{
		var fill = new Fill();
		var patternFill = new PatternFill { PatternType = PatternValues.Solid };
		patternFill.Append(new ForegroundColor { Rgb = GetHexBinaryValue(color) });
		fill.Append(patternFill);
		return fill;
	}

	private static Border CreateThinBorder(System.Drawing.Color color)
	{
		var border = new Border();
		border.Append(new LeftBorder { Color = GetColor(color), Style = BorderStyleValues.Thin });
		border.Append(new RightBorder { Color = GetColor(color), Style = BorderStyleValues.Thin });
		border.Append(new TopBorder { Color = GetColor(color), Style = BorderStyleValues.Thin });
		border.Append(new BottomBorder { Color = GetColor(color), Style = BorderStyleValues.Thin });
		return border;
	}

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

	private static uint GetNextConditionalNumberFormatId(Stylesheet stylesheet)
	{
		var nextNumberFormatId = stylesheet
			.Descendants<NumberingFormat>()
			.Select(nf => nf.NumberFormatId?.Value ?? 164U)
			.DefaultIfEmpty(164U)
			.Max() + 1;

		return Math.Max(nextNumberFormatId, 165U);
	}

	private static string EscapeFormulaText(string text)
		=> text.Replace("\"", "\"\"");
}
