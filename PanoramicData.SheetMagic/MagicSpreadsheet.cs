using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using PanoramicData.SheetMagic.Exceptions;
using PanoramicData.SheetMagic.Extensions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using Table = DocumentFormat.OpenXml.Spreadsheet.Table;

namespace PanoramicData.SheetMagic;

public class MagicSpreadsheet : IDisposable
{
	private const string Letters = "abcdefghijklmnopqrstuvwxyz";
	private const string Numbers = "0123456789";
	private static readonly Regex CellReferenceRegex = new(@"(?<col>([A-Z]|[a-z])+)(?<row>(\d)+)");

	private readonly FileInfo? _fileInfo;
	private readonly Stream? _stream;
	private readonly Options _options;
	private readonly HashSet<string> _uniqueTableDisplayNames = new();

	public MagicSpreadsheet(FileInfo fileInfo, Options options)
	{
		_fileInfo = fileInfo;
		_options = options;
	}

	public MagicSpreadsheet(FileInfo fileInfo)
		: this(fileInfo, new Options())
	{
	}

	public MagicSpreadsheet(Stream stream, Options options)
	{
		_stream = stream;
		_options = options;
	}

	public MagicSpreadsheet(Stream stream)
		: this(stream, new Options())
	{
	}

	private SpreadsheetDocument? _document;

	public List<string> SheetNames
		=> ((((_document ?? throw new InvalidOperationException("No document loaded."))
			.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart not created"))
			.Workbook ?? throw new InvalidOperationException("Workbook not created"))
			.Sheets ?? throw new InvalidOperationException("Sheets not created"))
			.ChildElements
			.Cast<Sheet>()
			.Select(s => s.Name?.Value ?? string.Empty)
			.ToList();

	public void Load() => _document = _fileInfo is not null
		? SpreadsheetDocument.Open(_fileInfo.FullName, false)
		: SpreadsheetDocument.Open(_stream!, false);

	private static string ColumnLetter(int intCol)
	{
		var intFirstLetter = (intCol / 676) + 64;
		var intSecondLetter = (intCol % 676 / 26) + 64;
		var intThirdLetter = (intCol % 26) + 65;

		var firstLetter = intFirstLetter > 64
			 ? (char)intFirstLetter
			 : ' ';
		var secondLetter = intSecondLetter > 64
			 ? (char)intSecondLetter
			 : ' ';
		var thirdLetter = (char)intThirdLetter;

		return string.Concat(firstLetter, secondLetter,
			 thirdLetter).Trim();
	}

	private static Cell CreateTextCell(string header, uint index, string text) =>
		 new(new InlineString(new Text { Text = text }))
		 {
			 DataType = CellValues.InlineString,
			 CellReference = header + index
		 };

	public void AddSheet<T>(List<T> items)
		=> AddSheet(items, null);


	public void AddSheet<T>(
		List<T> items,
		string? sheetName)
		=> AddSheet(items, sheetName, _options.DefaultAddSheetOptions.Clone());

	public void AddSheet<T>(
		List<T> items,
		string? sheetName,
		AddSheetOptions addSheetOptions)
	{
		if (items is null)
		{
			throw new ArgumentNullException(nameof(items));
		}

		// Were any items provided?
		if (items.Count == 0)
		{
			// No.  This is not permitted.

			// How should we fail?
			if (addSheetOptions.ThrowExceptionOnEmptyList)
			{
				throw new InvalidOperationException(
					"It is not permitted to add a sheet containing no items, as this would result in a corrupted XLSX file.  " +
					"To avoid this error, send an AddSheetOptions to the AddSheet call with ThrowExceptionOnEmptyList set to false.");
			}
			else
			{
				// Silently fail
				return;
			}
		}

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

		var type = typeof(T);
		var isExtended = type.IsGenericType && type.GetGenericTypeDefinition().UnderlyingSystemType.FullName == "PanoramicData.SheetMagic.Extended`1";
		var typeName = isExtended ? type.GenericTypeArguments[0].Name : type.Name;

		// Create a document for writing if not already loaded or created
		if (_document == null)
		{
			// Check that either a stream of document was provided
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
			if (_options.TableStyles.Count > 0)
			{
				var workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>("rId3");
				GenerateWorkbookStylesPart1Content(workbookStylesPart, _options.TableStyles[0]);
			}
		}

		// Set a sheet name if not provided
		const int SheetNameCharacterLimit = 31;
		if (string.IsNullOrWhiteSpace(sheetName))
		{
			try
			{
				// Get the length and leave space for an "s" on the end
				var length = Math.Min(typeName.Length, SheetNameCharacterLimit - 1);
				sheetName = $"{typeName.Substring(0, length)}s";
			}
			catch (Exception)
			{
				sheetName = "Sheet";
			}
		}

		// Fail if the sheetName is longer than the 31 character limit in Excel
		if (sheetName!.Length > SheetNameCharacterLimit)
		{
			throw new ArgumentException($"Sheet name cannot be more than {SheetNameCharacterLimit} characters", nameof(sheetName));
		}

		// Fail if there any sheets existing with the new sheet's name
		if (_document.WorkbookPart.Workbook.Sheets?.Any(existingSheet => string.Equals(((Sheet)existingSheet).Name.Value, sheetName, StringComparison.InvariantCultureIgnoreCase)) ?? false)
		{
			throw new ArgumentException($"Sheet name {sheetName} already exists. Sheet names must be unique per Workbook", nameof(sheetName));
		}

		var worksheetPart = CreateWorksheetPart(_document, sheetName);
		var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()
			?? throw new InvalidOperationException("No sheetdata in Worksheet.");

		// Determine property list
		var propertyList = new List<PropertyInfo>();
		var keyHashset = new HashSet<string>();
		var columnConfigurations = new Columns();
		Type basicType;
		if (isExtended)
		{
			basicType = type.GenericTypeArguments[0];
			var propertyInfo = type.GetProperties().Single(p => p.Name == nameof(Extended<object>.Properties));
			foreach (var item in items)
			{
				if (item is null)
				{
					continue;
				}

				var dictionary = (Dictionary<string, object>)propertyInfo.GetValue(item);

				var keys = dictionary.Keys.ToList();

				// Include/exclude as appropriate
				if (addSheetOptions.IncludeProperties?.Any() ?? false)
				{
					keys = keys.Where(key => addSheetOptions.IncludeProperties.Contains(key, StringComparer.InvariantCultureIgnoreCase)).ToList();
				}
				else if (addSheetOptions.ExcludeProperties?.Any() ?? false)
				{
					keys = keys.Where(key => !addSheetOptions.ExcludeProperties.Contains(key, StringComparer.InvariantCultureIgnoreCase)).ToList();
				}

				foreach (var key in keys)
				{
					keyHashset.Add(key);
				}
			}
		}
		else
		{
			basicType = type;
		}

		propertyList.AddRange(basicType.GetProperties());

		// Filter the propertyList according to the AddSheetOptions
		if (addSheetOptions.IncludeProperties?.Any() ?? false)
		{
			propertyList = propertyList.Where(p => addSheetOptions.IncludeProperties.Contains(p.Name, StringComparer.InvariantCultureIgnoreCase)).ToList();
		}
		else if (addSheetOptions.ExcludeProperties?.Any() ?? false)
		{
			propertyList = propertyList.Where(p => !addSheetOptions.ExcludeProperties.Contains(p.Name, StringComparer.InvariantCultureIgnoreCase)).ToList();
		}

		// By default, apply a sort
		var keyList = addSheetOptions.SortExtendedProperties
			? keyHashset.OrderBy(k => k).ToList()
			: keyHashset.ToList();

		// Add the columns
		var totalColumnCount = (uint)(propertyList.Count + keyList.Count);
		for (var n = 0; n < totalColumnCount; n++)
		{
			columnConfigurations.AppendChild(new Column
			{
				BestFit = true
			});
		}

		// Add header
		uint rowIndex = 0;
		var row = new Row { RowIndex = ++rowIndex };
		sheetData.AppendChild(row);
		var cellIndex = 0;

		foreach (var header in propertyList.Select(p => p.GetPropertyDescription() ?? p.Name))
		{
			row.AppendChild(CreateTextCell(
				ColumnLetter(cellIndex++),
				rowIndex,
				header ?? string.Empty)
				);
		}

		foreach (var header in keyList)
		{
			row.AppendChild(CreateTextCell(
				ColumnLetter(cellIndex++),
				rowIndex,
				header ?? string.Empty));
		}

		// Add sheet data
		var enumerableCellOptions = addSheetOptions?.EnumerableCellOptions ?? _options.EnumerableCellOptions;
		foreach (var item in items)
		{
			cellIndex = 0;
			row = new Row { RowIndex = ++rowIndex };
			sheetData.AppendChild(row);

			// Add cells for the properties
			foreach (var property in propertyList)
			{
				Cell? cell = null;
				if (isExtended)
				{
					var baseItem = type.GetProperties().Single(p => p.Name == nameof(Extended<object>.Item)).GetValue(item);
					cell = GetCell(
						enumerableCellOptions,
						property.GetValue(baseItem),
						cellIndex,
						rowIndex);
				}
				else
				{
					cell = GetCell(
						enumerableCellOptions,
						property.GetValue(item),
						cellIndex,
						rowIndex);
				}

				if (cell is not null)
				{
					row.AppendChild(cell);
				}

				cellIndex++;
			}

			// If not extended, this list will be empty
			if (isExtended)
			{
				var propertyInfo = type.GetProperties().Single(p => p.Name == nameof(Extended<object>.Properties));
				var dictionary = (Dictionary<string, object>)propertyInfo.GetValue(item);

				foreach (var key in keyList)
				{
					if (!dictionary.TryGetValue(key, out var @object))
					{
						@object = string.Empty;
					}

					// Don't add cells for null objects
					if (@object is not null)
					{
						var cell = CreateTextCell(ColumnLetter(cellIndex), rowIndex, @object.ToString());
						row.AppendChild(cell);
					}

					cellIndex++;
				}
			}
		}

		// Adding table style?
		if (addSheetOptions?.TableOptions == null)
		{
			return;
		}
		// Yes - apply style

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

		// Create the table
		var reference = $"A1:{ColumnLetter(columnConfigurations.Count() - 1)}{items.Count + 1}";
		var tableDefinitionPart = worksheetPart.AddNewPart<TableDefinitionPart>();
		tableDefinitionPart.Table = new Table
		{
			Id = (uint)_document.WorkbookPart.Workbook.Sheets.Count(),
			Name = addSheetOptions.TableOptions.Name,
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

		// Add the TableParts to the worksheet
		var tableParts = new TableParts
		{
			Count = 1U
		};
		tableParts.Append(new TablePart { Id = worksheetPart.GetIdOfPart(tableDefinitionPart) });
		worksheetPart.Worksheet.Append(tableParts);
	}

	private static Cell? GetCell<T>(
		EnumerableCellOptions enumerableCellOptions,
		T? v,
		int cellIndex,
		uint rowIndex)
	{
		object? value;
		if (
			enumerableCellOptions.Expand
			&& v is not string
			&& v is IEnumerable iEnumerable
			)
		{
			var stringBuilder = new StringBuilder();
			var isFirst = true;
			foreach (var il in iEnumerable)
			{
				if (!isFirst)
				{
					stringBuilder.Append(enumerableCellOptions.CellDelimiter);
				}

				stringBuilder.Append(il?.ToString() ?? "NULL");
				isFirst = false;
			}

			value = stringBuilder.ToString();
		}
		else
		{
			value = v?.ToString() ?? string.Empty;
		}

		var cell = CreateTextCell(
			ColumnLetter(cellIndex),
			rowIndex,
			value.ToString());
		return cell;
	}

	private static WorksheetPart CreateWorksheetPart(SpreadsheetDocument document, string? sheetName)
	{
		var worksheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>();
		var worksheet = new Worksheet(new SheetData());
		var sheet = new Sheet
		{
			Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
			SheetId = (uint)document.WorkbookPart.WorksheetParts.Count() + 1,
			Name = sheetName
		};
		document.WorkbookPart.Workbook.Sheets ??= new Sheets();
		document.WorkbookPart.Workbook.Sheets.AppendChild(sheet);

		worksheetPart.Worksheet = worksheet;
		return worksheetPart;
	}

	public void Save()
	{
		// Ensure that at least one sheet has been added
		if (_document?.WorkbookPart?.Workbook?.Sheets == null || !_document.WorkbookPart.Workbook.Sheets.Any())
		{
			// This has to contain some data to prevent file corruption.
			AddSheet(new[] { new { Error = "No data was output." } }.ToList(), "Sheet1");
		}

		if (_document?.WorkbookPart?.Workbook is null)
		{
			throw new InvalidOperationException("Document incorrectly created.");
		}

		_document.WorkbookPart.Workbook.Save();
		_document.Close();
	}

	private void ReleaseUnmanagedResources() => _document?.Dispose();

	public void Dispose()
	{
		ReleaseUnmanagedResources();
		GC.SuppressFinalize(this);
	}

	~MagicSpreadsheet()
	{
		ReleaseUnmanagedResources();
	}

	public List<T?> GetList<T>() where T : class, new()
		=> GetList<T>(null);

	public List<T?> GetList<T>(string? sheetName) where T : class, new()
		 => GetExtendedList<T>(sheetName).ConvertAll(e => e.Item);

	public List<Extended<T>> GetExtendedList<T>() where T : class, new()
		=> GetExtendedList<T>(null);

	/// <summary>
	/// Get sheet data
	/// </summary>
	/// <typeparam name="T">The type of object to load</typeparam>
	/// <param name="sheetName">The sheet name (if null, uses the first sheet in the workbook)</param>
	public List<Extended<T>> GetExtendedList<T>(string? sheetName) where T : class, new()
	{
		if (_document == null)
		{
			throw new InvalidOperationException("Document not loaded.");
		}

		Sheet sheet;
		var sheets = (_document.WorkbookPart ?? throw new InvalidOperationException("No WorkbookPart in document"))
			.Workbook
			.Sheets
			.Cast<Sheet>()
			.ToList();
		if (sheets.Count == 1)
		{
			sheet = sheets.Single();
		}
		else if (sheetName == null)
		{
			var typeName = typeof(T).Name;
			sheet = sheets.Find(s => StringsMatch(
				(s.Name
					?? throw new InvalidOperationException("Sheet contains no name")).Value
					?? throw new InvalidOperationException("Sheet name is null"),
				typeName)
			);
			if (sheet == null)
			{
				throw new ArgumentException($"Could not find sheet with a name matching type {typeName}.  Try specifying the name explicitly.  Available options {string.Join(", ", sheets.Select(s => s.Name))}");
			}
		}
		else
		{
			sheet = sheets.SingleOrDefault(s => s.Name == sheetName);
			if (sheet == null)
			{
				throw new ArgumentException($"Could not find sheet '{sheetName}'.  Available options {string.Join(", ", sheets.Select(s => s.Name))}");
			}
		}

		if (_document.WorkbookPart.GetPartById(sheet.Id.Value) is not WorksheetPart worksheetPart)
		{
			throw new FormatException($"No WorksheetPart found for sheet {sheet.Name}");
		}

		var worksheet = worksheetPart.Worksheet;
		// We have a worksheet part and a worksheet

		var stringTable = _document
			 .WorkbookPart
			 .GetPartsOfType<SharedStringTablePart>()
			 .FirstOrDefault();

		// Get the SheetData
		var sheetData = worksheet.GetFirstChild<SheetData>();

		// How many table parts are there on this worksheet part?
		var tableDefinitionParts = worksheetPart.TableDefinitionParts.ToList();
		if (tableDefinitionParts.Count > 1)
		{
			throw new FormatException($"Too many tables present on sheet {sheet.Name}.  Only one or zero may be present.");
		}

		// Get the rows and columns by
		List<Row> rows;
		List<string>? columns;

		var tableColumnOffset = 0;

		// Is there a table definition?
		var sheetDataRows = sheetData.Descendants<Row>();
		if (tableDefinitionParts.Count == 0)
		{
			// No - just use all rows and all columns
			rows = sheetDataRows.ToList();
			columns = rows.FirstOrDefault()?.Descendants<Cell>().Select(c => GetCellValueString(c, stringTable)).ToList();
		}
		else
		{
			// Yes - determine the rows and columns to use based on the table definition
			var table = tableDefinitionParts.Single().Table;
			var tableReference = table.Reference.Value;
			var tableReferenceArray = tableReference.Split(':');
			var fromReference = tableReferenceArray[0];
			var toReference = tableReferenceArray[1];

			var fromReferenceValue = GetReference(fromReference);
			var toReferenceValue = GetReference(toReference);

			var firstRowIndex = fromReferenceValue.rowIndex;
			var lastRowIndex = toReferenceValue.rowIndex;
			var firstColumnIndex = fromReferenceValue.columnIndex;
			var lastColumnIndex = toReferenceValue.columnIndex;
			tableColumnOffset = firstColumnIndex;

			rows = sheetDataRows
				 .SkipWhile(r => int.Parse(r.RowIndex) < firstRowIndex + 1)
				 .TakeWhile(r => int.Parse(r.RowIndex) <= lastRowIndex + 1)
				 .ToList();
			columns = rows
				 .FirstOrDefault()
				 ?.Descendants<Cell>()
				 .SkipWhile(c => GetReference(c.CellReference).columnIndex < firstColumnIndex)
				 .TakeWhile(c => GetReference(c.CellReference).columnIndex <= lastColumnIndex)
				 .Select(c => GetCellValueString(c, stringTable))
				 .ToList();
		}

		// Make sure that the columns match the type properties

		var tMappings = new Dictionary<int, string>();
		var extensionMappings = new Dictionary<int, string>();
		var columnIndex = 0;
		var properties = typeof(T).GetProperties().ToList();

		if (columns == null)
		{
			throw new FormatException("No columns found");
		}

		foreach (var column in columns)
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
					// OK, so a few fuzzy match, but do any completely match?
					var completelyMatchingProperty = properties.SingleOrDefault(p => string.Equals(p.Name, column, StringComparison.InvariantCultureIgnoreCase));
					if (completelyMatchingProperty != null)
					{
						tMappings[columnIndex] = completelyMatchingProperty.Name;
						break;
					}

					var matchingPropertiesString = string.Join("; ", matchingProperties.Select(p => p.Name));
					throw new FormatException($"More than one column matches {column}: {matchingPropertiesString}");
			}

			columnIndex++;
		}

		if (tMappings.Count != properties.Count && !_options.IgnoreUnmappedProperties)
		{
			var missingProperties = string.Join(", ", properties.Where(p => tMappings.Values.All(k => k != p.Name)));
			throw new InvalidOperationException($"Not all properties are mapped.  Missing: {missingProperties}");
		}

		// Process all rows UNTIL there is a row will no associated cells
		var list = new List<Extended<T>>();
		var rowIndex = 0;
		foreach (var row in rows.Skip(1))
		{
			rowIndex++;
			var cells = row.Descendants<Cell>().ToList();
			// Is the row empty?
			if (cells.All(cell => (GetCellValueDirect(cell, stringTable)?.ToString() ?? string.Empty)?.Length == 0))
			{
				// Yes.
				// Should we stop processing the table/sheet at this row.
				if (_options.StopProcessingOnFirstEmptyRow)
				{
					break;
				}

				// Can we just add null to the list?
				if (_options.EmptyRowInterpretedAsNull)
				{
					// Yes. Add an empty Extended<T> and move on to the next row.
					list.Add(new Extended<T>(default));
					continue;
				}
				// Unhandled empty row
				throw new EmptyRowException(rowIndex);
			}

			var item = new T();
			var eiProperties = new Dictionary<string, object?>();
			for (columnIndex = 0; columnIndex < columns.Count; columnIndex++)
			{
				var propertyName = tMappings.ContainsKey(columnIndex)
					 ? tMappings[columnIndex]
					 : extensionMappings[columnIndex];
				try
				{
					var property = properties.SingleOrDefault(p => p.Name == propertyName);
					var index = columnIndex;
					var cell = cells.SingleOrDefault(c => GetReference(c.CellReference.Value).columnIndex == index + tableColumnOffset);
					if (cell == null)
					{
						if (!_options.LoadNullExtendedProperties)
						{
							// No such cell.  Skip.
							continue;
						}

						throw new InvalidOperationException($"Null cell found for column '{propertyName}'");
					}

					if (property == null)
					{
						eiProperties[columns[columnIndex]] = GetCellValueDirect(cell, stringTable);
						continue;
					}
					// We have a property

					var propertyTypeName = property.PropertyType.IsGenericType
						 ? $"{property.PropertyType.GetGenericTypeDefinition().Name}<{string.Join(", ", property.PropertyType.GenericTypeArguments.Select(t => t.Name))}>"
						 : property.PropertyType.Name;
					switch (propertyTypeName)
					{
						case "Double":
							var cellValueDoubleObject = GetCellValue<double>(cell, stringTable);
							if (cellValueDoubleObject != null)
							{
								SetItemProperty(item, Convert.ToDouble(cellValueDoubleObject), propertyName);
							}

							break;
						case "Single":
							var cellValueFloatObject = GetCellValue<float>(cell, stringTable);
							if (cellValueFloatObject != null)
							{
								SetItemProperty(item, Convert.ToSingle(cellValueFloatObject), propertyName);
							}

							break;
						case "Int16":
							var cellValueShortObject = GetCellValue<short>(cell, stringTable);
							if (cellValueShortObject != null)
							{
								SetItemProperty(item, Convert.ToInt16(cellValueShortObject), propertyName);
							}

							break;
						case "UInt16":
							var cellValueUShortObject = GetCellValue<ushort>(cell, stringTable);
							if (cellValueUShortObject != null)
							{
								SetItemProperty(item, Convert.ToInt16(cellValueUShortObject), propertyName);
							}

							break;
						case "Int32":
							var cellValueIntObject = GetCellValue<int>(cell, stringTable);
							if (cellValueIntObject != null)
							{
								SetItemProperty(item, Convert.ToInt32(cellValueIntObject), propertyName);
							}

							break;
						case "UInt32":
							var cellValueUIntObject = GetCellValue<uint>(cell, stringTable);
							if (cellValueUIntObject != null)
							{
								SetItemProperty(item, Convert.ToUInt32(cellValueUIntObject), propertyName);
							}

							break;
						case "Int64":
							var cellValueLongObject = GetCellValue<long>(cell, stringTable);
							if (cellValueLongObject != null)
							{
								SetItemProperty(item, Convert.ToInt64(cellValueLongObject), propertyName);
							}

							break;
						case "UInt64":
							var cellValueULongObject = GetCellValue<ulong>(cell, stringTable);
							if (cellValueULongObject != null)
							{
								SetItemProperty(item, Convert.ToUInt64(cellValueULongObject), propertyName);
							}

							break;
						case "DateTime":
							var cellValueDateTimeObject = GetCellValue<DateTime>(cell, stringTable);
							if (cellValueDateTimeObject != null)
							{
								SetItemProperty(item, Convert.ToDateTime(cellValueDateTimeObject), propertyName);
							}

							break;
						case "String":
							SetItemProperty(item, (string?)GetCellValue<string>(cell, stringTable), propertyName);
							break;
						case "List`1<String>":
							var text = (string?)GetCellValue<string>(cell, stringTable);
							if (text is null)
							{
								SetItemProperty(item, new List<string>(), propertyName);
							}
							else
							{
								var stringList = text.Split(new[] { _options.ListSeparator }, StringSplitOptions.RemoveEmptyEntries).ToList();
								SetItemProperty(item, stringList, propertyName);
							}

							break;
						case "Nullable`1<Boolean>":
							{
								switch (GetCellValue<object?>(cell, stringTable))
								{
									case string stringValue2:
										if (string.IsNullOrWhiteSpace(stringValue2))
										{
											SetItemProperty(item, (bool?)null, propertyName);
										}
										else
										{
											SetItemProperty(item, (bool?)bool.Parse(stringValue2), propertyName);
										}

										break;
									case bool typedValue2:
										SetItemProperty(item, (bool?)typedValue2, propertyName);
										break;
									case null:
										SetItemProperty(item, null, propertyName);
										break;
									default:
										throw new InvalidOperationException("Invalid value type");
								}

								break;
							}
						case "Nullable`1<Double>":
							{
								switch (GetCellValue<object?>(cell, stringTable))
								{
									case string stringValue2:
										if (string.IsNullOrWhiteSpace(stringValue2))
										{
											SetItemProperty(item, (double?)null, propertyName);
										}
										else
										{
											SetItemProperty(item, (double?)double.Parse(stringValue2), propertyName);
										}

										break;
									case double typedValue2:
										SetItemProperty(item, (double?)typedValue2, propertyName);
										break;
									case null:
										SetItemProperty(item, null, propertyName);
										break;
									default:
										throw new InvalidOperationException("Invalid value type");
								}

								break;
							}
						case "Nullable`1<Single>":
							{
								switch (GetCellValue<object?>(cell, stringTable))
								{
									case string stringValue2:
										if (string.IsNullOrWhiteSpace(stringValue2))
										{
											SetItemProperty(item, (float?)null, propertyName);
										}
										else
										{
											SetItemProperty(item, (float?)float.Parse(stringValue2), propertyName);
										}

										break;
									case float typedValue2:
										SetItemProperty(item, (float?)typedValue2, propertyName);
										break;
									case null:
										SetItemProperty(item, null, propertyName);
										break;
									default:
										throw new InvalidOperationException("Invalid value type");
								}

								break;
							}
						case "Nullable`1<Int64>":
							{
								switch (GetCellValue<object?>(cell, stringTable))
								{
									case string stringValue2:
										if (string.IsNullOrWhiteSpace(stringValue2))
										{
											SetItemProperty(item, (long?)null, propertyName);
										}
										else
										{
											SetItemProperty(item, (long?)long.Parse(stringValue2), propertyName);
										}

										break;
									case long typedValue2:
										SetItemProperty(item, (long?)typedValue2, propertyName);
										break;
									case null:
										SetItemProperty(item, null, propertyName);
										break;
									default:
										throw new InvalidOperationException("Invalid value type");
								}

								break;
							}
						case "Nullable`1<Int32>":
							{
								switch (GetCellValue<object?>(cell, stringTable))
								{
									case string stringValue2:
										if (string.IsNullOrWhiteSpace(stringValue2))
										{
											SetItemProperty(item, (int?)null, propertyName);
										}
										else
										{
											SetItemProperty(item, (int?)int.Parse(stringValue2), propertyName);
										}

										break;
									case int typedValue2:
										SetItemProperty(item, (int?)typedValue2, propertyName);
										break;
									case null:
										SetItemProperty(item, null, propertyName);
										break;
									default:
										throw new InvalidOperationException("Invalid value type");
								}

								break;
							}
						case "Nullable`1<Int16>":
							{
								switch (GetCellValue<object?>(cell, stringTable))
								{
									case string stringValue2:
										if (string.IsNullOrWhiteSpace(stringValue2))
										{
											SetItemProperty(item, (short?)null, propertyName);
										}
										else
										{
											SetItemProperty(item, (short?)short.Parse(stringValue2), propertyName);
										}

										break;
									case short typedValue2:
										SetItemProperty(item, (short?)typedValue2, propertyName);
										break;
									case null:
										SetItemProperty(item, null, propertyName);
										break;
									default:
										throw new InvalidOperationException("Invalid value type");
								}

								break;
							}
						case "Nullable`1<DateTime>":
							{
								switch (GetCellValue<object?>(cell, stringTable))
								{
									case string stringValue2:
										if (string.IsNullOrWhiteSpace(stringValue2))
										{
											SetItemProperty(item, (DateTime?)null, propertyName);
										}
										else
										{
											SetItemProperty(item, (DateTime?)DateTime.Parse(stringValue2), propertyName);
										}

										break;
									case DateTime typedValue2:
										SetItemProperty(item, (DateTime?)typedValue2, propertyName);
										break;
									case null:
										SetItemProperty(item, null, propertyName);
										break;
									default:
										throw new InvalidOperationException("Invalid value type");
								}

								break;
							}
						default:
							// Is it an enum?
							var stringValue = (string?)GetCellValue<string>(cell, stringTable);
							if (property.PropertyType.IsEnum)
							{
								SetItemProperty(item, Enum.Parse(property.PropertyType, stringValue, true), propertyName);
							}
							else
							{
								throw new NotSupportedException($"Column index {columnIndex} matching {propertyName} has unsupported field type {propertyTypeName}.");
							}

							break;
					}
				}
				catch (Exception exception)
				{
					throw new ValidationException($"Issue with property '{propertyName}' on row {rowIndex}: {exception.Message}", exception);
				}
			}

			list.Add(new Extended<T>(item, eiProperties));
		}

		return list;
	}

	private static bool StringsMatch(string string1, string string2) => TweakString(string1) == TweakString(string2);

	internal static string TweakString(string text)
	{
		var stringBuilder = new StringBuilder();

		foreach (var @char in text.ToLowerInvariant())
		{
			if (!Letters.Contains(@char) && !Numbers.Contains(@char))
			{
				continue;
			}

			stringBuilder.Append(@char);
		}

		var tweakString = stringBuilder.ToString();

		// Chop numbers from the beginning
		while (tweakString.Length > 0 && Numbers.Contains(tweakString[0]))
		{
			tweakString = tweakString.Substring(1);
		}

		// Remove plurals
		return tweakString.EndsWith("s") && !tweakString.EndsWith("ss")
			 ? tweakString.Substring(0, tweakString.Length - 1)
			 : tweakString;
	}

	private string GetCellValueString(Cell cell, SharedStringTablePart stringTable)
	{
		var cellValueText = cell.CellValue?.Text;

		return cell.DataType != null && (CellValues)cell.DataType == CellValues.SharedString
			? stringTable.SharedStringTable.ElementAt(int.Parse(cellValueText)).InnerText
			: cellValueText ?? cell.InnerText;
	}

	private string? FormatCellAsNumber(Cell cell, string formatString)
	{
		// Check if it's a number
		if (double.TryParse(
			(cell.CellValue != null &&
			!string.IsNullOrEmpty(cell.CellValue.Text))
			? cell.CellValue?.Text
			: cell.InnerText, out var number))
		{
			return number.ToString(formatString);
		}

		return null;
	}

	private string? FormatCellAsDateTime(Cell cell, string formatString)
	{
		// Excel stores dates as a number (number of days since January 1, 1900),
		//so "44166" text is 03/12/2020
		// IF the number is an integer, it's only days. If it's a double, it's a fractional
		// portion of a day.
		var baseDate = new DateTime(1900, 01, 01);

		if (int.TryParse(
			(cell.CellValue != null &&
			!string.IsNullOrEmpty(cell.CellValue.Text))
			? cell.CellValue.Text
			: cell.InnerText, out var intDaysSinceBaseDate))
		{
			// See: https://www.kirix.com/stratablog/excel-date-conversion-days-from-1900
			// Note you DO have to take off 2 days!
			DateTime? actualDate = baseDate.AddDays(intDaysSinceBaseDate).AddDays(-2);

			// Return the date - we have to replace lower-case 'm' with upper-case as
			// required by C# else we get minutes
			// Some custom formats used by customers also have @ and ; in them.
			return actualDate.Value.ToString(
				formatString
					.Replace("\\", string.Empty)
					.Replace(";", string.Empty)
					.Replace("@", string.Empty)
					.Replace("m", "M"))
				.Trim();
		}

		// Could not parse cell value as an integer
		return null;
	}

	/// <summary>
	/// Is the format string a date string?
	/// </summary>
	/// <param name="formatString"></param>
	/// <returns></returns>
	private bool IsFormatStringADate(string formatString) =>
		formatString.IndexOf("d", StringComparison.OrdinalIgnoreCase) >= 0 ||
		formatString.IndexOf("m", StringComparison.OrdinalIgnoreCase) >= 0 ||
		formatString.IndexOf("y", StringComparison.OrdinalIgnoreCase) >= 0;

	private string? GetCellFormatFromStyle(Cell cell)
	{
		// NOTES:
		// First of all - when you choose a cell number format in Excel, you can choose from any of the 164 custom styles(not ALL are shown in the UI but they do exist).
		// You CANNOT retrieve ANY of the built -in number formats from an XML document.
		// You CANNOT guarantee that even if you could, that they are the same between different versions of Office, and between different locales.
		// So, for this to work, you HAVE TO choose a custom style in Excel and set it to what you want.
		// BUT, if you just choose any of the 'custom' formats, they are also actually built in -so if you see one in the list 'dd/mm/yyyy' and you select and OK it, it does NOT add a Custom Number Format to the XML document.
		// What you HAVE to do is enter a custom format and ensure that it is NOT in any of the lists. The easiest way to do this is to in this case type 'dd/mm/yyyy ' with a space, with no quotes, and then OK it.
		// This DOES then add it to the XML document and my code can read it, by finding the custom number format, it's style etc, etc.
		// HOWEVER, the internal representation of dates and times is days(and fractions of days) since Jan 1, 1900.
		// So the ONLY way I can tell if it's intended to be a date is by checking the format string (which I got from the XML style) and if it's anything like dd/mm/yyyy or dd-mmm-yyyy etc or any such combination that makes sense(and that is a manual to to see if the format string contains d / dd / ddd etc, then I can parse it and add the correct number of days and milliseconds onto the date and then format it as a string.
		// Annoyingly also, Excel date format is very flexible and can confuse with time format, e.g.you can put mm for months or for minutes (it treats upper case and lower case the same, it does some parsing to work out what you mean), so that in c# I also have to convert the lower case 'm' to upper case (as required by DateTime.Format).
		//This also means that this doesn't work at the moment for a time string. ONLY dates.

		// When check the cell.DataType, IF it is NULL then don't return null but first do the next steps (and if they fail / can't find then return null)
		// From the Cell element, check if it has a style Index. Get that (e.g. 14U)
		// Find , in Styles = > CellFormats (by index - starts at 0) the style by index from last step
		// If the style has a Number Format ID element (e.g. 167U) then ...
		// Find the Styles => NumberingFormats node where its NumberFormatId is the ID above
		// Then get the FormatCode attribute (e.g. "FormatCode = "dd/mm/yyyy;@" )
		// Then do treat it as text but do a ToString with the format above

		try
		{
			if (cell.StyleIndex?.HasValue == true)
			{
				var styleIndex = (int)cell.StyleIndex.Value;
				var cellFormats = _document?.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats;
				var numberingFormats = _document?.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats;

				if (cellFormats == null)
				{
					// Results in a string
					return null;
				}

				var cellFormat = (CellFormat)cellFormats.ElementAt(styleIndex);
				var formatString = string.Empty;

				if (cellFormat.NumberFormatId?.HasValue == true)
				{
					var numberFormatId = (uint)cellFormat.NumberFormatId.Value;

					if (numberFormatId >= BuiltInCellFormats.CustomFormatStartIndex)
					{
						// CUSTOM cell format
						var numberingFormat =
							numberingFormats
							.Cast<NumberingFormat>()
							.SingleOrDefault(f => f.NumberFormatId.Value == numberFormatId);

						if (numberingFormat == null)
						{
							return null;
						}

						// Get the format
						formatString = numberingFormat.FormatCode.Value;
					}
					else
					{
						// BUILT-IN number format
						var builtInFormat = BuiltInCellFormats.GetBuiltInCellFormatByIndex((int)numberFormatId);

						if (builtInFormat != null)
						{
							if (builtInFormat.Value.formatType == CellFormatType.General ||
								builtInFormat.Value.formatType == CellFormatType.Text ||
								builtInFormat.Value.formatType == CellFormatType.Unknown)
							{
								// Results in a string
								return null;
							}

							formatString = builtInFormat.Value.formatString;
						}
					}

					return IsFormatStringADate(formatString)
						? FormatCellAsDateTime(cell, formatString)
						: FormatCellAsNumber(cell, formatString);
				}
			}

			// Everything else (string)
			return null;
		}
		catch
		{
			// Results in a string value
			return null;
		}
	}

	private object? GetCellValueDirect(Cell cell, SharedStringTablePart stringTable)
	{
		var cellValueText = cell.CellValue?.Text;
		if (cell.DataType == null)
		{
			// Check whether there is a built-in style set
			if (GetCellFormatFromStyle(cell) is string text)
			{
				return text;
			}

			return cellValueText;
		}

		return (CellValues)cell.DataType switch
		{
			CellValues.SharedString => stringTable.SharedStringTable
									.ElementAt(int.Parse(cellValueText)).InnerText,
			CellValues.Boolean => cellValueText switch
			{
				"1" => true,
				"0" => false,
				_ => null,
			},
			CellValues.Number => double.Parse(cellValueText),
			CellValues.Date => DateTime.Parse(cellValueText),
			CellValues.Error or CellValues.String or CellValues.InlineString => cellValueText ?? cell.InnerText,
			_ => throw new NotSupportedException($"Unsupported data type {cell.DataType?.Value.ToString() ?? "None"}"),
		};
	}

	private object? GetCellValue<T>(Cell cell, SharedStringTablePart stringTable)
	{
		var cellValueText = cell.CellValue?.Text;
		if (cell.DataType == null)
		{
			switch (typeof(T).Name)
			{
				case "Int32":
					if (int.TryParse(cellValueText, out var intValue))
					{
						return intValue;
					}

					throw new FormatException($"Could not convert cell {cell.CellReference} to an integer.");
				case "Int64":
					if (long.TryParse(cellValueText, out var longValue))
					{
						return longValue;
					}

					throw new FormatException($"Could not convert cell {cell.CellReference} to an integer.");
				case "Double":
					if (int.TryParse(cellValueText, out var doubleValue))
					{
						return doubleValue;
					}

					throw new FormatException($"Could not convert cell {cell.CellReference} to a double.");
				case "Single":
					if (float.TryParse(cellValueText, out var floatValue))
					{
						return floatValue;
					}

					throw new FormatException($"Could not convert cell {cell.CellReference} to a double.");
				case "String":
				case "Object":
					return cellValueText;
				default:
					throw new NotSupportedException($"Unsupported data type {typeof(T).Name}");
			}
		}

		switch (cell.DataType is null ? null : (CellValues?)cell.DataType)
		{
			case null:
				return null;
			case CellValues.SharedString:
				var stringTableIndex = int.Parse(cellValueText);
				var sharedStringElement = stringTable.SharedStringTable.ElementAt(stringTableIndex);
				return sharedStringElement.InnerText;
			case CellValues.Boolean:
				return cellValueText switch
				{
					"1" => true,
					"0" => false,
					_ => null,
				};
			case CellValues.Number:
				return double.Parse(cellValueText);
			case CellValues.Date:
				return DateTime.Parse(cellValueText);
			case CellValues.Error:
			case CellValues.String:
			case CellValues.InlineString:
				try
				{
					return (T)Convert.ChangeType(cellValueText ?? cell.InnerText, typeof(T));
				}
				catch (FormatException)
				{
					return null;
				}
			default:
				throw new NotSupportedException($"Unsupported data type {cell.DataType?.Value.ToString() ?? "None"}");
		}
	}

	private (int columnIndex, int rowIndex) GetReference(string cellReference)
	{
		var match = CellReferenceRegex.Match(cellReference);

		if (match == null)
		{
			throw new ArgumentException($"Invalid cell reference {cellReference}", nameof(cellReference));
		}

		var col = match.Groups["col"].Value;
		var row = match.Groups["row"].Value;

		return (ExcelColumnNameToNumber(col) - 1, int.Parse(row) - 1);
	}

	private static int ExcelColumnNameToNumber(string columnName)
	{
		if (string.IsNullOrEmpty(columnName))
		{
			throw new ArgumentNullException(nameof(columnName));
		}

		var sum = 0;
		foreach (var t in columnName.ToUpperInvariant())
		{
			sum *= 26;
			sum += t - 'A' + 1;
		}

		return sum;
	}

	private void SetItemProperty<T, T1>(T item, T1 cellValue, string propertyName)
	{
		var cellValues = new List<object?> { cellValue };
		typeof(T).InvokeMember(propertyName,
			 BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty,
			 Type.DefaultBinder, item, cellValues.ToArray());
	}

	private void SetItemProperty<T>(T item, object? cellValue, string propertyName)
	{
		var cellValues = new List<object?> { cellValue };
		typeof(T).InvokeMember(propertyName,
			 BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty,
			 Type.DefaultBinder, item, cellValues.ToArray());
	}

	private void GenerateWorkbookStylesPart1Content(WorkbookStylesPart workbookStylesPart1, CustomTableStyle customTableStyle)
	{
		var stylesheet1 = new Stylesheet { MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "x14ac x16r2 xr xr9" } };
		stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
		stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
		stylesheet1.AddNamespaceDeclaration("x16r2", "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main");
		stylesheet1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
		stylesheet1.AddNamespaceDeclaration("xr9", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision9");

		// Fonts
		var fonts = new Fonts { Count = 1U, KnownFonts = true };
		var font = new Font();
		font.Append(new FontSize { Val = 11D });
		font.Append(new Color { Theme = 1U });
		font.Append(new FontName { Val = "Calibri" });
		font.Append(new FontFamilyNumbering { Val = 2 });
		font.Append(new FontScheme { Val = FontSchemeValues.Minor });
		fonts.Append(font);

		// Fills
		var fills = new Fills { Count = 2U };
		var noneFill = new Fill();
		noneFill.Append(new PatternFill { PatternType = PatternValues.None });
		var gray125Fill = new Fill();
		gray125Fill.Append(new PatternFill { PatternType = PatternValues.Gray125 });
		fills.Append(noneFill);
		fills.Append(gray125Fill);

		// Outer Borders
		var borders = new Borders { Count = 1U };
		var outerBorder = new Border();
		outerBorder.Append(new LeftBorder());
		outerBorder.Append(new RightBorder());
		outerBorder.Append(new TopBorder());
		outerBorder.Append(new BottomBorder());
		outerBorder.Append(new DiagonalBorder());
		borders.Append(outerBorder);

		var cellStyleFormats1 = new CellStyleFormats { Count = 1U };
		var cellFormat1 = new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U };

		cellStyleFormats1.Append(cellFormat1);

		var cellFormats1 = new CellFormats { Count = 1U };
		var cellFormat2 = new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U };

		cellFormats1.Append(cellFormat2);

		var cellStyles1 = new CellStyles { Count = 1U };
		var cellStyle1 = new CellStyle { Name = "Normal", FormatId = 0U, BuiltinId = 0U };

		cellStyles1.Append(cellStyle1);

		var differentialFormats = new DifferentialFormats { Count = 3U };
		var tableStyles1 = new TableStyles { Count = 1U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

		var tableStyleCount = 0U;
		if (customTableStyle.OddRowStyle != null)
		{
			tableStyleCount++;
		}

		if (customTableStyle.EvenRowStyle != null)
		{
			tableStyleCount++;
		}

		if (customTableStyle.HeaderRowStyle != null)
		{
			tableStyleCount++;
		}

		if (customTableStyle.WholeTableStyle != null)
		{
			tableStyleCount++;
		}

		var tableStyle1 = new TableStyle { Name = customTableStyle.Name, Pivot = false, Count = tableStyleCount };
		tableStyle1.SetAttribute(new OpenXmlAttribute("xr9", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision9", "{640A183E-9F4E-4A71-80D9-2176963C18AB}"));
		tableStyles1.Append(tableStyle1);
		var tableStyleIndex = 0U;
		AddTableStyleElement(customTableStyle.OddRowStyle, differentialFormats, tableStyle1, tableStyleIndex++, TableStyleValues.FirstRowStripe);
		AddTableStyleElement(customTableStyle.EvenRowStyle, differentialFormats, tableStyle1, tableStyleIndex++, TableStyleValues.SecondRowStripe);
		AddTableStyleElement(customTableStyle.HeaderRowStyle, differentialFormats, tableStyle1, tableStyleIndex++, TableStyleValues.HeaderRow);
		AddTableStyleElement(customTableStyle.WholeTableStyle, differentialFormats, tableStyle1, tableStyleIndex, TableStyleValues.WholeTable);

		// Colors
		var colors1 = new Colors();

		var mruColors1 = new MruColors();
		var color5 = new Color { Rgb = "FFE1CCF0" };

		mruColors1.Append(color5);

		colors1.Append(mruColors1);

		stylesheet1.Append(fonts);
		stylesheet1.Append(fills);
		stylesheet1.Append(borders);
		stylesheet1.Append(cellStyleFormats1);
		stylesheet1.Append(cellFormats1);
		stylesheet1.Append(cellStyles1);
		stylesheet1.Append(differentialFormats);
		stylesheet1.Append(tableStyles1);
		stylesheet1.Append(colors1);

		workbookStylesPart1.Stylesheet = stylesheet1;
	}

	private void AddTableStyleElement(
		TableRowStyle? thisCustomTableStyle,
		DifferentialFormats differentialFormats,
		TableStyle tableStyle1,
		uint tableStyleIndex,
		TableStyleValues tableStyleValues)
	{
		if (thisCustomTableStyle is null)
		{
			return;
		}

		var differentialFormat = new DifferentialFormat();

		// Font color
		if (thisCustomTableStyle.FontColor.HasValue)
		{
			var font = new Font();
			if (thisCustomTableStyle.FontWeight == FontWeight.Bold)
			{
				font.Append(new Bold());
			}

			font.Append(GetColor(thisCustomTableStyle.FontColor.Value));
			differentialFormat.Append(font);
		}

		// Background color
		if (thisCustomTableStyle.BackgroundColor.HasValue)
		{
			var fill = new Fill();
			var patternFill = new PatternFill();
			patternFill.Append(new BackgroundColor { Rgb = GetHexBinaryValue(thisCustomTableStyle.BackgroundColor.Value) });
			fill.Append(patternFill);
			differentialFormat.Append(fill);
		}

		// Inner border
		if (thisCustomTableStyle.InnerBorderColor.HasValue || thisCustomTableStyle.OuterBorderColor.HasValue)
		{
			var border = new Border();

			if (thisCustomTableStyle.OuterBorderColor.HasValue)
			{
				border.Append(new LeftBorder { Color = GetColor(thisCustomTableStyle.OuterBorderColor.Value), Style = BorderStyleValues.Thin });
				border.Append(new RightBorder { Color = GetColor(thisCustomTableStyle.OuterBorderColor.Value), Style = BorderStyleValues.Thin });
				border.Append(new TopBorder { Color = GetColor(thisCustomTableStyle.OuterBorderColor.Value), Style = BorderStyleValues.Thin });
				border.Append(new BottomBorder { Color = GetColor(thisCustomTableStyle.OuterBorderColor.Value), Style = BorderStyleValues.Thin });
			}

			if (thisCustomTableStyle.InnerBorderColor.HasValue)
			{
				border.Append(new VerticalBorder { Color = GetColor(thisCustomTableStyle.InnerBorderColor.Value), Style = BorderStyleValues.Thin });
				border.Append(new HorizontalBorder { Color = GetColor(thisCustomTableStyle.InnerBorderColor.Value), Style = BorderStyleValues.Thin });
			}

			differentialFormat ??= new DifferentialFormat();
			differentialFormat.Append(border);
		}

		differentialFormats.Append(differentialFormat);
		tableStyle1.Append(new TableStyleElement { Type = tableStyleValues, FormatId = tableStyleIndex });
	}

	private Color GetColor(System.Drawing.Color color)
		=> Equals(color, System.Drawing.Color.White)
			? new Color { Theme = 0U }
			: new Color { Rgb = GetHexBinaryValue(color) };

	private static HexBinaryValue GetHexBinaryValue(System.Drawing.Color color) => new()
	{
		Value = $"FF{color.R:X2}{color.G:X2}{color.B:X2}"
	};
}