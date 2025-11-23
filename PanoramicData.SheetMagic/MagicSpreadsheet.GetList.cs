using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using PanoramicData.SheetMagic.Exceptions;
using PanoramicData.SheetMagic.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PanoramicData.SheetMagic;

/// <summary>
/// Methods for reading sheets from the workbook - Part 1: Main entry points and helpers
/// </summary>
public partial class MagicSpreadsheet
{
	public List<T?> GetList<T>() where T : class, new()
		=> GetList<T>(null);

	public List<T?> GetList<T>(string? sheetName) where T : class, new()
		 => GetExtendedList<T>(sheetName).ConvertAll(static e => e.Item);

	public List<Extended<T>> GetExtendedList<T>() where T : class, new()
		=> GetExtendedList<T>(null);

	/// <summary>
	/// Get sheet data
	/// </summary>
	/// <typeparam name="T">The type of object to load</typeparam>
	/// <param name="sheetName">The sheet name (if null, uses the first sheet in the workbook)</param>
	public List<Extended<T>> GetExtendedList<T>(string? sheetName) where T : class, new()
	{
		ValidateDocumentLoaded();

		var sheet = FindSheet<T>(sheetName);
		var worksheetPart = GetWorksheetPart(sheet);
		var stringTable = GetStringTable();
		var sheetData = GetSheetData(sheet, worksheetPart);

		ValidateTableDefinitionParts(sheet, worksheetPart);

		var (rows, columns, tableColumnOffset) = GetRowsAndColumns(worksheetPart, sheetData, stringTable);
		var (tMappings, extensionMappings) = MapColumnsToProperties<T>(columns);

		ValidatePropertyMappings<T>(tMappings);

		return ProcessDataRows<T>(rows, columns!, tMappings, extensionMappings, tableColumnOffset, stringTable);
	}

	private void ValidateDocumentLoaded()
	{
		if (_document == null)
		{
			throw new InvalidOperationException("Document not loaded.");
		}
	}

	private Sheet FindSheet<T>(string? sheetName)
	{
		var sheets = (_document!.WorkbookPart ?? throw new InvalidOperationException("No WorkbookPart in document"))
			.Workbook
			.Sheets
			?.Cast<Sheet>()
			.ToList()
			?? throw new InvalidOperationException("No Sheets in Workbook");

		if (sheets.Count == 1)
		{
			return sheets.Single();
		}

		if (sheetName == null)
		{
			return FindSheetByTypeName<T>(sheets);
		}

		return sheets.SingleOrDefault(s => s.Name == sheetName)
			?? throw new ArgumentException($"Could not find sheet '{sheetName}'.  Available options {string.Join(", ", sheets.Select(s => s.Name))}");
	}

	private static Sheet FindSheetByTypeName<T>(List<Sheet> sheets)
	{
		var typeName = typeof(T).Name;
		return sheets.Find(s => StringsMatch(
			(s.Name
				?? throw new InvalidOperationException("Sheet contains no name")).Value
				?? throw new InvalidOperationException("Sheet name is null"),
			typeName)
		) ?? throw new ArgumentException($"Could not find sheet with a name matching type {typeName}.  Try specifying the name explicitly.  Available options {string.Join(", ", sheets.Select(s => s.Name))}");
	}

	private WorksheetPart GetWorksheetPart(Sheet sheet)
		=> _document!.WorkbookPart!.GetPartById(sheet.Id!.Value!) as WorksheetPart
			?? throw new FormatException($"No WorksheetPart found for sheet {sheet.Name}");

	private SharedStringTablePart? GetStringTable()
		=> _document!.WorkbookPart!.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

	private static SheetData GetSheetData(Sheet sheet, WorksheetPart worksheetPart)
		=> worksheetPart.Worksheet.GetFirstChild<SheetData>()
			?? throw new FormatException($"No SheetData found for workSheet {sheet.Name}");

	private static void ValidateTableDefinitionParts(Sheet sheet, WorksheetPart worksheetPart)
	{
		var tableDefinitionParts = worksheetPart.TableDefinitionParts.ToList();
		if (tableDefinitionParts.Count > 1)
		{
			throw new FormatException($"Too many tables present on sheet {sheet.Name}.  Only one or zero may be present.");
		}
	}

	private static (List<Row> rows, List<string>? columns, int tableColumnOffset) GetRowsAndColumns(
		WorksheetPart worksheetPart,
		SheetData sheetData,
		SharedStringTablePart? stringTable)
	{
		var tableDefinitionParts = worksheetPart.TableDefinitionParts.ToList();
		var sheetDataRows = sheetData.Descendants<Row>();

		if (tableDefinitionParts.Count == 0)
		{
			return GetRowsAndColumnsWithoutTable(sheetDataRows, stringTable);
		}
		else
		{
			return GetRowsAndColumnsFromTable(tableDefinitionParts, sheetDataRows, stringTable);
		}
	}

	private static (List<Row> rows, List<string>? columns, int tableColumnOffset) GetRowsAndColumnsWithoutTable(
		IEnumerable<Row> sheetDataRows,
		SharedStringTablePart? stringTable)
	{
		var rows = sheetDataRows.ToList();
		var columns = rows.FirstOrDefault()?.Descendants<Cell>().Select(c => GetCellValueString(c, stringTable)).ToList();
		return (rows, columns, 0);
	}

	private static (List<Row> rows, List<string>? columns, int tableColumnOffset) GetRowsAndColumnsFromTable(
		List<TableDefinitionPart> tableDefinitionParts,
		IEnumerable<Row> sheetDataRows,
		SharedStringTablePart? stringTable)
	{
		var table = tableDefinitionParts.Single().Table;
		var tableReference = table.Reference!.Value!;
		var tableReferenceArray = tableReference.Split(':');

		var fromReference = tableReferenceArray[0];
		var toReference = tableReferenceArray[1];

		var fromReferenceValue = GetReference(fromReference);
		var toReferenceValue = GetReference(toReference);

		var firstRowIndex = fromReferenceValue.rowIndex;
		var lastRowIndex = toReferenceValue.rowIndex;
		var firstColumnIndex = fromReferenceValue.columnIndex;
		var lastColumnIndex = toReferenceValue.columnIndex;

		var rows = sheetDataRows
			.SkipWhile(r => int.Parse(r.RowIndex!) < firstRowIndex + 1)
			.TakeWhile(r => int.Parse(r.RowIndex!) <= lastRowIndex + 1)
			.ToList();

		var columns = rows
			.FirstOrDefault()
			?.Descendants<Cell>()
			.SkipWhile(c => GetReference(c.CellReference!).columnIndex < firstColumnIndex)
			.TakeWhile(c => GetReference(c.CellReference!).columnIndex <= lastColumnIndex)
			.Select(c => GetCellValueString(c, stringTable))
			.ToList();

		return (rows, columns, firstColumnIndex);
	}

	private static (Dictionary<int, string> tMappings, Dictionary<int, string> extensionMappings) MapColumnsToProperties<T>(List<string>? columns)
	{
		if (columns == null)
		{
			throw new FormatException("No columns found");
		}

		var tMappings = new Dictionary<int, string>();
		var extensionMappings = new Dictionary<int, string>();
		var properties = typeof(T).GetProperties().ToList();

		for (var columnIndex = 0; columnIndex < columns.Count; columnIndex++)
		{
			MapColumn(columns[columnIndex], columnIndex, properties, tMappings, extensionMappings);
		}

		return (tMappings, extensionMappings);
	}

	private static void MapColumn(
		string column,
		int columnIndex,
		List<System.Reflection.PropertyInfo> properties,
		Dictionary<int, string> tMappings,
		Dictionary<int, string> extensionMappings)
	{
		var matchingProperties = properties
			.Where(p => StringsMatch(p.GetPropertyDescription() ?? p.Name, column))
			.ToList();

		switch (matchingProperties.Count)
		{
			case 0:
				extensionMappings[columnIndex] = column;
				break;
			case 1:
				tMappings[columnIndex] = matchingProperties.Single().Name;
				break;
			default:
				HandleMultipleMatches(column, columnIndex, properties, tMappings, matchingProperties);
				break;
		}
	}

	private static void HandleMultipleMatches(
		string column,
		int columnIndex,
		List<System.Reflection.PropertyInfo> properties,
		Dictionary<int, string> tMappings,
		List<System.Reflection.PropertyInfo> matchingProperties)
	{
		var completelyMatchingProperty = properties.SingleOrDefault(p => string.Equals(p.Name, column, StringComparison.InvariantCultureIgnoreCase));

		if (completelyMatchingProperty != null)
		{
			tMappings[columnIndex] = completelyMatchingProperty.Name;
		}
		else
		{
			var matchingPropertiesString = string.Join("; ", matchingProperties.Select(p => p.Name));
			throw new FormatException($"More than one column matches {column}: {matchingPropertiesString}");
		}
	}

	private void ValidatePropertyMappings<T>(Dictionary<int, string> tMappings)
	{
		var properties = typeof(T).GetProperties().ToList();

		if (tMappings.Count != properties.Count && !_options.IgnoreUnmappedProperties)
		{
			var missingProperties = string.Join(", ", properties.Where(p => tMappings.Values.All(k => k != p.Name)));
			throw new InvalidOperationException($"Not all properties are mapped.  Missing: {missingProperties}");
		}
	}

	private List<Extended<T>> ProcessDataRows<T>(
		List<Row> rows,
		List<string> columns,
		Dictionary<int, string> tMappings,
		Dictionary<int, string> extensionMappings,
		int tableColumnOffset,
		SharedStringTablePart? stringTable) where T : class, new()
	{
		var list = new List<Extended<T>>();
		var rowIndex = 0;
		var properties = typeof(T).GetProperties().ToList();

		foreach (var row in rows.Skip(1))
		{
			rowIndex++;
			var cells = row.Descendants<Cell>().ToList();

			if (ShouldSkipEmptyRow(cells, rowIndex, stringTable))
			{
				if (_options.StopProcessingOnFirstEmptyRow)
				{
					break;
				}

				if (_options.EmptyRowInterpretedAsNull)
				{
					list.Add(new Extended<T>(default));
					continue;
				}

				throw new EmptyRowException(rowIndex);
			}

			var (item, eiProperties) = ProcessRow<T>(cells, columns, tMappings, extensionMappings, properties, tableColumnOffset, rowIndex, stringTable);
			list.Add(new Extended<T>(item, eiProperties));
		}

		return list;
	}

	private bool ShouldSkipEmptyRow(List<Cell> cells, int rowIndex, SharedStringTablePart? stringTable)
		=> cells.All(cell => (GetCellValueDirect(cell, stringTable)?.ToString() ?? string.Empty)?.Length == 0);

	private (T item, Dictionary<string, object?> eiProperties) ProcessRow<T>(
		List<Cell> cells,
		List<string> columns,
		Dictionary<int, string> tMappings,
		Dictionary<int, string> extensionMappings,
		List<System.Reflection.PropertyInfo> properties,
		int tableColumnOffset,
		int rowIndex,
		SharedStringTablePart? stringTable) where T : class, new()
	{
		var item = new T();
		var eiProperties = new Dictionary<string, object?>();

		for (var columnIndex = 0; columnIndex < columns.Count; columnIndex++)
		{
			var propertyName = tMappings.TryGetValue(columnIndex, out var value)
				? value
				: extensionMappings[columnIndex];

			try
			{
				ProcessColumnValue(item, cells, columns, properties, propertyName, columnIndex, tableColumnOffset, eiProperties, stringTable);
			}
			catch (Exception exception)
			{
				throw new ValidationException($"Issue with property '{propertyName}' on row {rowIndex}: {exception.Message}", exception);
			}
		}

		return (item, eiProperties);
	}

	private void ProcessColumnValue<T>(
		T item,
		List<Cell> cells,
		List<string> columns,
		List<System.Reflection.PropertyInfo> properties,
		string propertyName,
		int columnIndex,
		int tableColumnOffset,
		Dictionary<string, object?> eiProperties,
		SharedStringTablePart? stringTable) where T : class, new()
	{
		var property = properties.SingleOrDefault(p => p.Name == propertyName);
		var cell = cells.SingleOrDefault(c => GetReference(c.CellReference!.Value!).columnIndex == columnIndex + tableColumnOffset);

		if (cell == null)
		{
			HandleMissingCell(property, columns, columnIndex, eiProperties, propertyName);
			return;
		}

		if (property == null)
		{
			eiProperties[columns[columnIndex]] = GetCellValueDirect(cell, stringTable);
		}
		else
		{
			ProcessCellValue(item, cell, property, propertyName, stringTable);
		}
	}

	private void HandleMissingCell(
		System.Reflection.PropertyInfo? property,
		List<string> columns,
		int columnIndex,
		Dictionary<string, object?> eiProperties,
		string propertyName)
	{
		if (property == null)
		{
			// This is an extended property - add empty string to maintain consistent property counts
			eiProperties[columns[columnIndex]] = string.Empty;
			return;
		}

		if (!_options.LoadNullExtendedProperties)
		{
			// No such cell.  Skip.
			return;
		}

		throw new InvalidOperationException($"Null cell found for column '{propertyName}'");
	}
}
