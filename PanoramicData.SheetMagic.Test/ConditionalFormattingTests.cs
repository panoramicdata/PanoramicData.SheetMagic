using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DrawingColor = System.Drawing.Color;

namespace PanoramicData.SheetMagic.Test;

public class ConditionalFormattingTests : Test
{
	[Fact]
	public void AddSheet_ConditionalFormattingWithMultipleRules_WritesRulesAndDifferentialFormats()
	{
		var fileInfo = GetXlsxTempFileInfo();

		try
		{
			var items = new List<ConditionalFormattingRow>
			{
				new("Alpha", "First", null),
				new("Bravo", "Second", 9)
			};

			var addSheetOptions = CreateAddSheetOptions(
				[nameof(ConditionalFormattingRow.Score)],
				CreateContainsBlanksRule(DrawingColor.Red),
				new ConditionalFormatRule
				{
					RuleType = ConditionalFormatRuleType.CellIs,
					Operator = ConditionalFormatOperator.GreaterThan,
					Formula = "5",
					Style = new ConditionalFormatStyle
					{
						FontColor = DrawingColor.Green
					}
				});

			using var document = WriteAndReopen(fileInfo, items, "Scores", addSheetOptions);
			var worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;
			var conditionalFormatting = Assert.Single(worksheet.Elements<ConditionalFormatting>());
			Assert.Equal("C2:C3", GetSqRef(conditionalFormatting));

			AssertBlanksThenGreaterThanRules(conditionalFormatting);
			AssertRedFillThenGreenFontFormats(document);
		}
		finally
		{
			fileInfo.Delete();
		}
	}

	[Fact]
	public void AddSheet_ConditionalFormattingForMultipleColumns_WritesOneBlockPerColumn()
	{
		var fileInfo = GetXlsxTempFileInfo();

		try
		{
			var items = new List<ConditionalFormattingRow>
			{
				new(null, "First", 1),
				new("Bravo", null, 2)
			};

			var addSheetOptions = CreateAddSheetOptions(
				[nameof(ConditionalFormattingRow.Name), nameof(ConditionalFormattingRow.Description)],
				new ConditionalFormatRule
				{
					RuleType = ConditionalFormatRuleType.ContainsBlanks,
					Style = new ConditionalFormatStyle
					{
						BackgroundColor = DrawingColor.Red
					}
				});

			using var document = WriteAndReopen(fileInfo, items, "Columns", addSheetOptions);
			var conditionalFormattings = GetConditionalFormattingsInColumnOrder(document);

			Assert.Equal(2, conditionalFormattings.Count);
			Assert.Equal("A2:A3", GetSqRef(conditionalFormattings[0]));
			Assert.Equal("B2:B3", GetSqRef(conditionalFormattings[1]));
			Assert.Equal("LEN(TRIM(A2))=0", conditionalFormattings[0].Elements<ConditionalFormattingRule>().Single().Elements<Formula>().Single().Text);
			Assert.Equal("LEN(TRIM(B2))=0", conditionalFormattings[1].Elements<ConditionalFormattingRule>().Single().Elements<Formula>().Single().Text);
		}
		finally
		{
			fileInfo.Delete();
		}
	}

	[Fact]
	public void AddSheet_ConditionalFormattingWithoutColumnNames_AppliesToAllColumns()
	{
		var fileInfo = GetXlsxTempFileInfo();

		try
		{
			var items = new List<ConditionalFormattingRow>
			{
				new("Alpha", "First", 1),
				new("Bravo", "Second", 2)
			};

			var addSheetOptions = CreateAddSheetOptions(
				columnNames: null,
				new ConditionalFormatRule
				{
					RuleType = ConditionalFormatRuleType.ContainsErrors,
					Style = new ConditionalFormatStyle
					{
						FontWeight = FontWeight.Bold
					}
				});

			using var document = WriteAndReopen(fileInfo, items, "AllColumns", addSheetOptions);
			var conditionalFormattings = GetConditionalFormattingsInColumnOrder(document);

			Assert.Equal(["A2:A3", "B2:B3", "C2:C3"], conditionalFormattings.Select(GetSqRef).ToArray());
			Assert.Equal(["ISERROR(A2)", "ISERROR(B2)", "ISERROR(C2)"], conditionalFormattings
				.Select(cf => cf.Elements<ConditionalFormattingRule>().Single().Elements<Formula>().Single().Text)
				.ToArray());

			var dxfs = GetDifferentialFormats(document);
			Assert.Equal(3, dxfs.Count);
			Assert.All(dxfs, dxf => Assert.Single(dxf.Elements<Font>().Single().Elements<Bold>()));
		}
		finally
		{
			fileInfo.Delete();
		}
	}

	private static string GetSqRef(ConditionalFormatting conditionalFormatting)
		=> conditionalFormatting.GetAttribute("sqref", string.Empty).Value ?? string.Empty;

	private static ConditionalFormatRule CreateContainsBlanksRule(DrawingColor backgroundColor)
		=> new()
		{
			RuleType = ConditionalFormatRuleType.ContainsBlanks,
			Style = new ConditionalFormatStyle
			{
				BackgroundColor = backgroundColor
			}
		};

	private static void AssertBlanksThenGreaterThanRules(ConditionalFormatting conditionalFormatting)
	{
		var rules = conditionalFormatting.Elements<ConditionalFormattingRule>().ToList();
		Assert.Equal(2, rules.Count);

		Assert.Equal(ConditionalFormatValues.ContainsBlanks, rules[0].Type?.Value);
		Assert.Equal("LEN(TRIM(C2))=0", rules[0].Elements<Formula>().Single().Text);
		Assert.Equal(1, (int)rules[0].Priority!.Value);

		Assert.Equal(ConditionalFormatValues.CellIs, rules[1].Type?.Value);
		Assert.Equal(ConditionalFormattingOperatorValues.GreaterThan, rules[1].Operator?.Value);
		Assert.Equal("5", rules[1].Elements<Formula>().Single().Text);
		Assert.Equal(2, (int)rules[1].Priority!.Value);
	}

	private static void AssertRedFillThenGreenFontFormats(SpreadsheetDocument document)
	{
		var dxfs = GetDifferentialFormats(document);
		Assert.Equal(2, dxfs.Count);
		Assert.Equal("FFFF0000", dxfs[0].Descendants<ForegroundColor>().Single().Rgb?.Value);
		Assert.Equal("FF008000", dxfs[1].Descendants<DocumentFormat.OpenXml.Spreadsheet.Color>().Single(x => x.Rgb is not null).Rgb?.Value);
	}

	/// <summary>
	/// Writes the sheet, then reopens it so that the assertions run against what actually reached
	/// the file rather than against the in-memory document.
	/// </summary>
	private static SpreadsheetDocument WriteAndReopen(
		FileInfo fileInfo,
		List<ConditionalFormattingRow> items,
		string sheetName,
		AddSheetOptions addSheetOptions)
	{
		using (var spreadsheet = new MagicSpreadsheet(fileInfo))
		{
			spreadsheet.AddSheet(items, sheetName, addSheetOptions);
			spreadsheet.Save();
		}

		return SpreadsheetDocument.Open(fileInfo.FullName, false);
	}

	private static List<ConditionalFormatting> GetConditionalFormattingsInColumnOrder(SpreadsheetDocument document)
		=> [.. document.WorkbookPart!
			.WorksheetParts
			.Single()
			.Worksheet!
			.Elements<ConditionalFormatting>()
			.OrderBy(GetSqRef)];

	private static AddSheetOptions CreateAddSheetOptions(List<string>? columnNames, params ConditionalFormatRule[] rules)
		=> new()
		{
			ConditionalFormats =
			[
				new ConditionalFormat
				{
					ColumnNames = columnNames,
					Rules = [.. rules]
				}
			]
		};

	private static List<DifferentialFormat> GetDifferentialFormats(SpreadsheetDocument document)
	{
		var differentialFormats = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.GetFirstChild<DifferentialFormats>();
		Assert.NotNull(differentialFormats);
		return [.. differentialFormats!.Elements<DifferentialFormat>()];
	}

	private sealed record ConditionalFormattingRow(string? Name, string? Description, int? Score);
}