using PanoramicData.SheetMagic.Extensions;

namespace PanoramicData.SheetMagic;

/// <summary>
/// Table definition support for added sheets
/// </summary>
public partial class MagicSpreadsheet
{
	private void ApplyTableStyleIfRequested<T>(
		List<T> items,
		AddSheetOptions addSheetOptions,
		WorksheetPart worksheetPart,
		SheetLayout layout)
	{
		if (addSheetOptions?.TableOptions == null)
		{
			return;
		}

		var tableColumns = CreateTableColumns(layout);
		var tableDefinitionPart = CreateTableDefinition(items, addSheetOptions, worksheetPart, layout.ColumnConfigurations, tableColumns);
		AttachTableToWorksheet(worksheetPart, tableDefinitionPart);
	}

	private static TableColumns CreateTableColumns(SheetLayout layout)
	{
		var tableColumns = new TableColumns { Count = layout.TotalColumnCount };
		var columnIndex = 0;
		var combinedList = layout.PropertyList
			.Select(p => p.GetPropertyDescription() ?? p.Name)
			.Concat(layout.KeyList)
			.ToList();

		foreach (var columnConfiguration in layout.ColumnConfigurations)
		{
			tableColumns.Append(new TableColumn
			{
				Name = combinedList[columnIndex],
				Id = (uint)++columnIndex,
			});
		}

		return tableColumns;
	}

	private TableDefinitionPart CreateTableDefinition<T>(
		List<T> items,
		AddSheetOptions addSheetOptions,
		WorksheetPart worksheetPart,
		Columns columnConfigurations,
		TableColumns tableColumns)
	{
		var reference = $"A1:{ColumnLetter(columnConfigurations.Count() - 1)}{items.Count + 1}";
		var tableDefinitionPart = worksheetPart.AddNewPart<TableDefinitionPart>();
		tableDefinitionPart.Table = new Table
		{
			Id = (uint)(_document!.WorkbookPart!.Workbook!.Sheets?.Count() ?? 0),
			Name = addSheetOptions.TableOptions!.Name,
			DisplayName = addSheetOptions.TableOptions.DisplayName,
			Reference = reference,
			TotalsRowShown = addSheetOptions.TableOptions.ShowTotalsRow,
			AutoFilter = new AutoFilter { Reference = reference },
			TableColumns = tableColumns,
			TableStyleInfo = new TableStyleInfo
			{
				Name = addSheetOptions.TableOptions.CustomTableStyle ?? addSheetOptions.TableOptions.XlsxTableStyle.ToString(),
				ShowFirstColumn = addSheetOptions.TableOptions.ShowFirstColumn,
				ShowLastColumn = addSheetOptions.TableOptions.ShowLastColumn,
				ShowRowStripes = addSheetOptions.TableOptions.ShowRowStripes,
				ShowColumnStripes = addSheetOptions.TableOptions.ShowColumnStripes
			},
		};
		tableDefinitionPart.Table.Save();
		return tableDefinitionPart;
	}

	private static void AttachTableToWorksheet(WorksheetPart worksheetPart, TableDefinitionPart tableDefinitionPart)
	{
		var tableParts = new TableParts { Count = 1U };
		tableParts.Append(new TablePart { Id = worksheetPart.GetIdOfPart(tableDefinitionPart) });
		worksheetPart.Worksheet!.Append(tableParts);
	}
}
