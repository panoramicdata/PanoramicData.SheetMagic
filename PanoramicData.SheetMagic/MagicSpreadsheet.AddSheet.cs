using DocumentFormat.OpenXml;
using PanoramicData.SheetMagic.Extensions;
using Table = DocumentFormat.OpenXml.Spreadsheet.Table;

namespace PanoramicData.SheetMagic;

/// <summary>
/// Methods for adding sheets to the workbook
/// </summary>
public partial class MagicSpreadsheet
{
	private const int SheetNameCharacterLimit = 31;

	/// <summary>
	/// Adds a sheet with the specified items using default options.
	/// </summary>
	/// <typeparam name="T">The type of items to add.</typeparam>
	/// <param name="items">The list of items to write to the sheet.</param>
	public void AddSheet<T>(List<T> items)
		=> AddSheet(items, null);

	/// <summary>
	/// Adds a sheet with the specified items and sheet name.
	/// </summary>
	/// <typeparam name="T">The type of items to add.</typeparam>
	/// <param name="items">The list of items to write to the sheet.</param>
	/// <param name="sheetName">The name of the sheet, or null to auto-generate.</param>
	public void AddSheet<T>(
		List<T> items,
		string? sheetName)
		=> AddSheet(items, sheetName, _options.DefaultAddSheetOptions.Clone());

	/// <summary>
	/// Adds a sheet with the specified items, sheet name, and options.
	/// </summary>
	/// <typeparam name="T">The type of items to add.</typeparam>
	/// <param name="items">The list of items to write to the sheet.</param>
	/// <param name="sheetName">The name of the sheet, or null to auto-generate.</param>
	/// <param name="addSheetOptions">Options for configuring how the sheet is added.</param>
	public void AddSheet<T>(
		List<T> items,
		string? sheetName,
		AddSheetOptions addSheetOptions)
	{
		ArgumentNullException.ThrowIfNull(items);

		if (!ValidateAndHandleEmptyItems(items, addSheetOptions))
		{
			return;
		}

		ValidateAndPrepareTableOptions(addSheetOptions);

		var (type, isExtended, isJObject, typeName) = AnalyzeTypeInfo<T>();

		EnsureDocumentExists();
		sheetName = DetermineAndValidateSheetName(sheetName, typeName);

		var worksheetPart = CreateWorksheetPart(_document!, sheetName);
		var sheetData = GetSheetData(worksheetPart);

		var (propertyList, columnConfigurations, keyList, totalColumnCount) =
			PopulateSheetData(items, addSheetOptions, type, isExtended, isJObject, sheetData);

		ApplyTableStyleIfRequested(items, addSheetOptions, worksheetPart, propertyList, columnConfigurations, keyList, totalColumnCount);
	}

	private static bool ValidateAndHandleEmptyItems<T>(List<T> items, AddSheetOptions addSheetOptions)
	{
		if (items.Count == 0)
		{
			if (addSheetOptions.ThrowExceptionOnEmptyList)
			{
				throw new InvalidOperationException(
					"It is not permitted to add a sheet containing no items, as this would result in a corrupted XLSX file.  " +
					"To avoid this error, send an AddSheetOptions to the AddSheet call with ThrowExceptionOnEmptyList set to false.");
			}

			return false; // Silently fail
		}

		return true;
	}

	private void ValidateAndPrepareTableOptions(AddSheetOptions addSheetOptions)
	{
		if (addSheetOptions.TableOptions is not null)
		{
			if (_uniqueTableDisplayNames.Contains(addSheetOptions.TableOptions.DisplayName))
			{
				addSheetOptions.TableOptions.DisplayName = $"{addSheetOptions.TableOptions.DisplayName}_{_uniqueTableDisplayNames.Count}";
			}
		}

		addSheetOptions.Validate(_options.TableStyles);

		if (addSheetOptions.TableOptions?.DisplayName != null)
		{
			if (!_uniqueTableDisplayNames.Add(addSheetOptions.TableOptions.DisplayName))
			{
				throw new ArgumentException($"Table DisplayName must be unique. There is already a Table with the DisplayName {addSheetOptions.TableOptions.DisplayName}");
			}
		}
	}

	private static (Type type, bool isExtended, bool isJObject, string typeName) AnalyzeTypeInfo<T>()
	{
		var type = typeof(T);
		var isExtended = type.IsGenericType && type.GetGenericTypeDefinition().UnderlyingSystemType.FullName == "PanoramicData.SheetMagic.Extended`1";
		var isJObject = type.FullName == "Newtonsoft.Json.Linq.JObject";
		var typeName = isExtended ? type.GenericTypeArguments[0].Name : type.Name;
		return (type, isExtended, isJObject, typeName);
	}

	private void EnsureDocumentExists()
	{
		if (_document != null)
		{
			return;
		}

		if (_stream is not null)
		{
			_document = SpreadsheetDocument.Create(_stream, SpreadsheetDocumentType.Workbook);
		}
		else
		{
			if (_fileInfo is null)
			{
				throw new InvalidOperationException("No file or stream provided.");
			}

			_document = SpreadsheetDocument.Create(_fileInfo.FullName, SpreadsheetDocumentType.Workbook);
		}

		var workbookPart = _document.AddWorkbookPart();
		workbookPart.Workbook = new Workbook();

		// Add any custom table styles
		var workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>("rId3");
		GenerateWorkbookStylesPart1Content(workbookStylesPart);
	}

	private string DetermineAndValidateSheetName(string? sheetName, string typeName)
	{
		sheetName = DetermineSheetName(sheetName, typeName);
		ValidateSheetNameLength(sheetName);
		ValidateSheetNameUniqueness(sheetName);
		return sheetName;
	}

	private static string DetermineSheetName(string? sheetName, string typeName)
	{
		if (!string.IsNullOrWhiteSpace(sheetName))
		{
			return sheetName!;
		}

		try
		{
			// Get the length and leave space for an "s" on the end
			var length = Math.Min(typeName.Length, SheetNameCharacterLimit - 1);
			return $"{typeName[..length]}s";
		}
		catch (Exception)
		{
			return "Sheet";
		}
	}

	private static void ValidateSheetNameLength(string sheetName)
	{
		if (sheetName.Length > SheetNameCharacterLimit)
		{
			throw new ArgumentException($"Sheet name cannot be more than {SheetNameCharacterLimit} characters", nameof(sheetName));
		}
	}

	private void ValidateSheetNameUniqueness(string sheetName)
	{
		var sheetExists = _document!.WorkbookPart!.Workbook.Sheets?
			.Any(existingSheet => string.Equals(((Sheet)existingSheet).Name!.Value, sheetName, StringComparison.InvariantCultureIgnoreCase)) ?? false;

		if (sheetExists)
		{
			throw new ArgumentException($"Sheet name {sheetName} already exists. Sheet names must be unique per Workbook", nameof(sheetName));
		}
	}

	private static SheetData GetSheetData(WorksheetPart worksheetPart)
		=> worksheetPart.Worksheet.GetFirstChild<SheetData>()
			?? throw new InvalidOperationException("No SheetData in Worksheet.");

	private (List<PropertyInfo> propertyList, Columns columnConfigurations, List<string> keyList, uint totalColumnCount)
		PopulateSheetData<T>(
			List<T> items,
			AddSheetOptions addSheetOptions,
			Type type,
			bool isExtended,
			bool isJObject,
			SheetData sheetData)
	{
		if (!isJObject)
		{
			AddItems(
				items,
				addSheetOptions,
				type,
				isExtended,
				sheetData,
				out var propertyList,
				out var columnConfigurations,
				out var keyList,
				out var totalColumnCount
			);
			return (propertyList, columnConfigurations, keyList, totalColumnCount);
		}
		else
		{
			AddJObjectItems(
				items,
				addSheetOptions,
				sheetData,
				out var propertyList,
				out var columnConfigurations,
				out var keyList,
				out var totalColumnCount
			);
			return (propertyList, columnConfigurations, keyList, totalColumnCount);
		}
	}

	private void ApplyTableStyleIfRequested<T>(
		List<T> items,
		AddSheetOptions addSheetOptions,
		WorksheetPart worksheetPart,
		List<PropertyInfo> propertyList,
		Columns columnConfigurations,
		List<string> keyList,
		uint totalColumnCount)
	{
		if (addSheetOptions?.TableOptions == null)
		{
			return;
		}

		var tableColumns = CreateTableColumns(propertyList, keyList, totalColumnCount, columnConfigurations);
		var tableDefinitionPart = CreateTableDefinition(items, addSheetOptions, worksheetPart, columnConfigurations, tableColumns);
		AttachTableToWorksheet(worksheetPart, tableDefinitionPart);
	}

	private static TableColumns CreateTableColumns(
		List<PropertyInfo> propertyList,
		List<string> keyList,
		uint totalColumnCount,
		Columns columnConfigurations)
	{
		var tableColumns = new TableColumns { Count = totalColumnCount };
		var columnIndex = 0;
		var combinedList = propertyList
			.Select(p => p.GetPropertyDescription() ?? p.Name)
			.Concat(keyList)
			.ToList();

		foreach (var columnConfiguration in columnConfigurations)
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
			Id = (uint)(_document!.WorkbookPart!.Workbook.Sheets?.Count() ?? 0),
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
		worksheetPart.Worksheet.Append(tableParts);
	}

	private void AddJObjectItems<T>(
		List<T> items,
		AddSheetOptions addSheetOptions,
		SheetData sheetData,
		out List<PropertyInfo> propertyList,
		out Columns columnConfigurations,
		out List<string> keyList,
		out uint totalColumnCount)
		=> throw new NotImplementedException("JObjects not yet supported.  Use Extended<JObject> instead.");

	private static WorksheetPart CreateWorksheetPart(SpreadsheetDocument document, string? sheetName)
	{
		var worksheetPart = document.WorkbookPart!.AddNewPart<WorksheetPart>();
		var worksheet = new Worksheet(new SheetData());
		var sheet = new Sheet
		{
			Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
			SheetId = (uint)document.WorkbookPart.WorksheetParts.Count() + 1,
			Name = sheetName
		};
		document.WorkbookPart.Workbook.Sheets ??= new Sheets();
		_ = document.WorkbookPart.Workbook.Sheets.AppendChild(sheet);

		worksheetPart.Worksheet = worksheet;
		return worksheetPart;
	}
}
