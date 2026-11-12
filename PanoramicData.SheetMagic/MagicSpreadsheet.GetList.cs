using PanoramicData.SheetMagic.Extensions;

namespace PanoramicData.SheetMagic;

/// <summary>
/// Methods for reading sheets from the workbook - Part 1: Main entry points and helpers
/// </summary>
public partial class MagicSpreadsheet
{
	/// <summary>
	/// Gets a list of items of type T from the first matching sheet.
	/// </summary>
	/// <typeparam name="T">The type of object to load.</typeparam>
	/// <returns>A list of items loaded from the sheet.</returns>
	public List<T?> GetList<T>() where T : class, new()
		=> GetList<T>(null);

	/// <summary>
	/// Gets a list of items of type T from the specified sheet.
	/// </summary>
	/// <typeparam name="T">The type of object to load.</typeparam>
	/// <param name="sheetName">The sheet name, or null to use the first matching sheet.</param>
	/// <returns>A list of items loaded from the sheet.</returns>
	public List<T?> GetList<T>(string? sheetName) where T : class, new()
		 => GetExtendedList<T>(sheetName).ConvertAll(static e => e.Item);

	/// <summary>
	/// Gets an extended list with additional properties from the first matching sheet.
	/// </summary>
	/// <typeparam name="T">The type of object to load.</typeparam>
	/// <returns>A list of Extended items with properties not mapped to T.</returns>
	public List<Extended<T>> GetExtendedList<T>() where T : class, new()
		=> GetExtendedList<T>(null);

	/// <summary>
	/// Gets an extended list with additional properties from the specified sheet.
	/// </summary>
	/// <typeparam name="T">The type of object to load.</typeparam>
	/// <param name="sheetName">The sheet name, or null to use the first matching sheet.</param>
	/// <returns>A list of Extended items with properties not mapped to T.</returns>
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
			.Workbook!
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
		=> worksheetPart.Worksheet!.GetFirstChild<SheetData>()
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
		var table = tableDefinitionParts.Single().Table!;
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
}
