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
					var dxf = DifferentialFormatBuilder.CreateDifferentialFormatFromStyle(rule.Style, ref nextNumberFormatId);
					differentialFormats.Append(dxf);

					var cfRule = ConditionalFormatRuleBuilder.BuildConditionalFormattingRule(rule, currentDxfId, priority++, firstCellRef);
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

	private static uint GetNextConditionalNumberFormatId(Stylesheet stylesheet)
	{
		var nextNumberFormatId = stylesheet
			.Descendants<NumberingFormat>()
			.Select(nf => nf.NumberFormatId?.Value ?? 164U)
			.DefaultIfEmpty(164U)
			.Max() + 1;

		return Math.Max(nextNumberFormatId, 165U);
	}
}
