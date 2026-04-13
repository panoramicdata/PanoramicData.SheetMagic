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
				new() { Name = "Alpha", Description = "First", Score = null },
				new() { Name = "Bravo", Description = "Second", Score = 9 }
			};

			var addSheetOptions = new AddSheetOptions
			{
				ConditionalFormats =
				[
					new ConditionalFormat
					{
						ColumnNames = [nameof(ConditionalFormattingRow.Score)],
						Rules =
						[
							new ConditionalFormatRule
							{
								RuleType = ConditionalFormatRuleType.ContainsBlanks,
								Style = new ConditionalFormatStyle
								{
									BackgroundColor = DrawingColor.Red
								}
							},
							new ConditionalFormatRule
							{
								RuleType = ConditionalFormatRuleType.CellIs,
								Operator = ConditionalFormatOperator.GreaterThan,
								Formula = "5",
								Style = new ConditionalFormatStyle
								{
									FontColor = DrawingColor.Green
								}
							}
						]
					}
				]
			};

			using (var spreadsheet = new MagicSpreadsheet(fileInfo))
			{
				spreadsheet.AddSheet(items, "Scores", addSheetOptions);
				spreadsheet.Save();
			}

			using var document = SpreadsheetDocument.Open(fileInfo.FullName, false);
			var worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet;
			var conditionalFormatting = Assert.Single(worksheet.Elements<ConditionalFormatting>());
			Assert.Equal("C2:C3", GetSqRef(conditionalFormatting));

			var rules = conditionalFormatting.Elements<ConditionalFormattingRule>().ToList();
			Assert.Equal(2, rules.Count);
			Assert.Equal(ConditionalFormatValues.ContainsBlanks, rules[0].Type?.Value);
			Assert.Equal("LEN(TRIM(C2))=0", rules[0].Elements<Formula>().Single().Text);
			Assert.Equal(1, (int)rules[0].Priority!.Value);

			Assert.Equal(ConditionalFormatValues.CellIs, rules[1].Type?.Value);
			Assert.Equal(ConditionalFormattingOperatorValues.GreaterThan, rules[1].Operator?.Value);
			Assert.Equal("5", rules[1].Elements<Formula>().Single().Text);
			Assert.Equal(2, (int)rules[1].Priority!.Value);

			var differentialFormats = document.WorkbookPart.WorkbookStylesPart!.Stylesheet.GetFirstChild<DifferentialFormats>();
			Assert.NotNull(differentialFormats);

			var dxfs = differentialFormats!.Elements<DifferentialFormat>().ToList();
			Assert.Equal(2, dxfs.Count);
			Assert.Equal("FFFF0000", dxfs[0].Descendants<ForegroundColor>().Single().Rgb?.Value);
			Assert.Equal("FF008000", dxfs[1].Descendants<DocumentFormat.OpenXml.Spreadsheet.Color>().Single(x => x.Rgb is not null).Rgb?.Value);
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
				new() { Name = null, Description = "First", Score = 1 },
				new() { Name = "Bravo", Description = null, Score = 2 }
			};

			var addSheetOptions = new AddSheetOptions
			{
				ConditionalFormats =
				[
					new ConditionalFormat
					{
						ColumnNames = [nameof(ConditionalFormattingRow.Name), nameof(ConditionalFormattingRow.Description)],
						Rules =
						[
							new ConditionalFormatRule
							{
								RuleType = ConditionalFormatRuleType.ContainsBlanks,
								Style = new ConditionalFormatStyle
								{
									BackgroundColor = DrawingColor.Red
								}
							}
						]
					}
				]
			};

			using (var spreadsheet = new MagicSpreadsheet(fileInfo))
			{
				spreadsheet.AddSheet(items, "Columns", addSheetOptions);
				spreadsheet.Save();
			}

			using var document = SpreadsheetDocument.Open(fileInfo.FullName, false);
			var conditionalFormattings = document.WorkbookPart!
				.WorksheetParts
				.Single()
				.Worksheet
				.Elements<ConditionalFormatting>()
				.OrderBy(GetSqRef)
				.ToList();

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
				new() { Name = "Alpha", Description = "First", Score = 1 },
				new() { Name = "Bravo", Description = "Second", Score = 2 }
			};

			var addSheetOptions = new AddSheetOptions
			{
				ConditionalFormats =
				[
					new ConditionalFormat
					{
						Rules =
						[
							new ConditionalFormatRule
							{
								RuleType = ConditionalFormatRuleType.ContainsErrors,
								Style = new ConditionalFormatStyle
								{
									FontWeight = FontWeight.Bold
								}
							}
						]
					}
				]
			};

			using (var spreadsheet = new MagicSpreadsheet(fileInfo))
			{
				spreadsheet.AddSheet(items, "AllColumns", addSheetOptions);
				spreadsheet.Save();
			}

			using var document = SpreadsheetDocument.Open(fileInfo.FullName, false);
			var conditionalFormattings = document.WorkbookPart!
				.WorksheetParts
				.Single()
				.Worksheet
				.Elements<ConditionalFormatting>()
				.OrderBy(GetSqRef)
				.ToList();

			Assert.Equal(["A2:A3", "B2:B3", "C2:C3"], conditionalFormattings.Select(GetSqRef).ToArray());
			Assert.Equal(["ISERROR(A2)", "ISERROR(B2)", "ISERROR(C2)"], conditionalFormattings
				.Select(cf => cf.Elements<ConditionalFormattingRule>().Single().Elements<Formula>().Single().Text)
				.ToArray());

			var differentialFormats = document.WorkbookPart.WorkbookStylesPart!.Stylesheet.GetFirstChild<DifferentialFormats>();
			Assert.NotNull(differentialFormats);
			Assert.Equal(3, differentialFormats!.Elements<DifferentialFormat>().Count());
			Assert.All(differentialFormats.Elements<DifferentialFormat>(), dxf => Assert.Single(dxf.Elements<Font>().Single().Elements<Bold>()));
		}
		finally
		{
			fileInfo.Delete();
		}
	}

	private static string GetSqRef(ConditionalFormatting conditionalFormatting)
		=> conditionalFormatting.GetAttribute("sqref", string.Empty).Value ?? string.Empty;

	private sealed class ConditionalFormattingRow
	{
		public string? Name { get; set; }

		public string? Description { get; set; }

		public int? Score { get; set; }
	}
}