using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using PanoramicData.SheetMagic.Exceptions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace PanoramicData.SheetMagic
{
    public class MagicSpreadsheet : IDisposable
    {
        private const string Letters = "abcdefghijklmnopqrstuvwxyz";
        private const string Numbers = "0123456789";
        private static readonly Regex CellReferenceRegex = new Regex(@"(?<col>([A-Z]|[a-z])+)(?<row>(\d)+)");

        private readonly FileInfo _fileInfo;
        private readonly Options _options;

        public MagicSpreadsheet(FileInfo fileInfo, Options? options = default)
        {
            _fileInfo = fileInfo;
            _options = options ?? new Options();
        }

        private SpreadsheetDocument? _document;

        public List<string> SheetNames =>
            _document?
            .WorkbookPart
            .Workbook
            .Sheets
            .ChildElements
            .Cast<Sheet>()
            .Select(s => s.Name.Value)
            .ToList()
            ?? throw new InvalidOperationException("No document loaded.");

        public void Load() => _document = SpreadsheetDocument.Open(_fileInfo.FullName, false);

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

        private static Cell CreateTextCell(string header, uint index, string? text) =>
            new Cell(new InlineString(new Text { Text = text }))
            {
                DataType = CellValues.InlineString,
                CellReference = header + index
            };

        public void AddSheet<T>(List<T> items, string? sheetName = null, AddSheetOptions? addSheetOptions = null)
        {
            var theAddSheetOptions = addSheetOptions ?? _options.DefaultAddSheetOptions;
            theAddSheetOptions.Validate();

            var type = typeof(T);
            var isExtended = type.IsGenericType && type.GetGenericTypeDefinition().UnderlyingSystemType.FullName == "PanoramicData.SheetMagic.Extended`1";
            var typeName = isExtended ? type.GenericTypeArguments[0].Name : type.Name;

            // Create a document for writing if not already loaded or created
            if (_document == null)
            {
                _document = SpreadsheetDocument.Create(_fileInfo.FullName, SpreadsheetDocumentType.Workbook);
                _document.AddWorkbookPart().Workbook = new Workbook();
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

            WorksheetPart worksheetPart = CreateWorksheetPart(_document, sheetName);
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            // Determine property list
            var propertyList = new List<PropertyInfo>();
            var keyHashset = new HashSet<string>();
            var columnConfigurations = new Columns();
            Type basicType;
            if (isExtended)
            {
                basicType = type.GenericTypeArguments[0];
                PropertyInfo propertyInfo = type.GetProperties().Single(p => p.Name == nameof(Extended<object>.Properties));
                foreach (var item in items)
                {
                    if (item is null)
                    {
                        continue;
                    }
                    var dictionary = (Dictionary<string, object>)propertyInfo.GetValue(item);
                    foreach (var key in dictionary.Keys)
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
            if (theAddSheetOptions?.IncludeProperties?.Any() ?? false)
            {
                propertyList = propertyList.Where(p => theAddSheetOptions.IncludeProperties.Contains(p.Name, StringComparer.InvariantCultureIgnoreCase)).ToList();
            }
            else if (theAddSheetOptions?.ExcludeProperties?.Any() ?? false)
            {
                propertyList = propertyList.Where(p => !theAddSheetOptions.ExcludeProperties.Contains(p.Name, StringComparer.InvariantCultureIgnoreCase)).ToList();
            }

            var keyList = keyHashset.OrderBy(k => k).ToList();

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

            foreach (var header in propertyList.Select(p => p.Name))
            {
                row.AppendChild(CreateTextCell(ColumnLetter(cellIndex++),
                    rowIndex, header ?? string.Empty));
            }
            foreach (var header in keyList)
            {
                row.AppendChild(CreateTextCell(ColumnLetter(cellIndex++),
                    rowIndex, header ?? string.Empty));
            }

            // Add sheet data
            foreach (var item in items)
            {
                cellIndex = 0;
                row = new Row { RowIndex = ++rowIndex };
                sheetData.AppendChild(row);

                // Add cells for the properties
                foreach (var property in propertyList)
                {
                    Cell cell;
                    if (isExtended)
                    {
                        PropertyInfo propertyInfo = type.GetProperties().Single(p => p.Name == nameof(Extended<object>.Item));
                        var baseItem = propertyInfo.GetValue(item);
                        cell = CreateTextCell(ColumnLetter(cellIndex++), rowIndex, property.GetValue(baseItem)?.ToString() ?? string.Empty);
                    }
                    else
                    {
                        cell = CreateTextCell(ColumnLetter(cellIndex++), rowIndex, property.GetValue(item)?.ToString() ?? string.Empty);
                    }
                    row.AppendChild(cell);
                }

                // If not extended, this list will be empty
                if (isExtended)
                {
                    PropertyInfo propertyInfo = type.GetProperties().Single(p => p.Name == nameof(Extended<object>.Properties));
                    var dictionary = (Dictionary<string, object>)propertyInfo.GetValue(item);

                    foreach (var key in keyList)
                    {
                        if (!dictionary.TryGetValue(key, out var @object))
                        {
                            @object = string.Empty;
                        }
                        var cell = CreateTextCell(ColumnLetter(cellIndex++), rowIndex, @object?.ToString());
                        row.AppendChild(cell);
                    }
                }
            }

            // Adding table style?
            if (theAddSheetOptions?.TableOptions == null)
            {
                return;
            }
            // Yes - apply style

            var tableColumns = new TableColumns() { Count = totalColumnCount };
            var columnIndex = 0;
            var combinedList = propertyList.Select(p => p.Name).Concat(keyList).ToList();
            foreach (var columnConfiguration in columnConfigurations)
            {
                tableColumns.Append(new TableColumn
                {
                    Name = combinedList[columnIndex],
                    Id = (uint)++columnIndex,
                });
            }

            // Determine the range
            var reference = $"A1:{ColumnLetter(columnConfigurations.Count())}{items.Count + 1}";
            var tableDefinitionPart = worksheetPart.AddNewPart<TableDefinitionPart>("rId1");
            tableDefinitionPart.Table = new Table
            (
                new AutoFilter() { Reference = reference },
                tableColumns
                )
            {
                Id = 1U,
                Name = theAddSheetOptions.TableOptions.Name,
                DisplayName = theAddSheetOptions.TableOptions.DisplayName,
                Reference = reference,
                TotalsRowShown = theAddSheetOptions.TableOptions.TotalsRowShown,
                TableStyleInfo = new TableStyleInfo()
                {
                    Name = theAddSheetOptions.TableOptions.XlsxTableStyle.ToString(),
                    ShowFirstColumn = theAddSheetOptions.TableOptions.ShowFirstColumn,
                    ShowLastColumn = theAddSheetOptions.TableOptions.ShowLastColumn,
                    ShowRowStripes = theAddSheetOptions.TableOptions.ShowRowStripes,
                    ShowColumnStripes = theAddSheetOptions.TableOptions.ShowColumnStripes
                }
            };
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
            (document.WorkbookPart.Workbook.Sheets ?? (document.WorkbookPart.Workbook.Sheets = new Sheets())).AppendChild(sheet);

            worksheetPart.Worksheet = worksheet;
            return worksheetPart;
        }

        public void Save()
        {
            // Ensure that at least one sheet has been added
            if (_document?.WorkbookPart?.Workbook?.Sheets == null || !_document.WorkbookPart.Workbook.Sheets.Any())
            {
                AddSheet(new List<object>(), "Sheet1");
            }
            if (_document?.WorkbookPart?.Workbook is null)
            {
                throw new Exception("Document incorrectly created.");
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

        public List<T?> GetList<T>(string? sheetName = null) where T : class, new()
            => GetExtendedList<T>(sheetName).Select(e => e.Item).ToList();

        /// <summary>
        /// Get sheet data
        /// </summary>
        /// <typeparam name="T">The type of object to load</typeparam>
        /// <param name="sheetName">The sheet name (if null, uses the first sheet in the workbook)</param>
        public List<Extended<T>> GetExtendedList<T>(string? sheetName = null) where T : class, new()
        {
            if (_document == null)
            {
                throw new InvalidOperationException("Document not loaded.");
            }

            Sheet sheet;
            var sheets = _document.WorkbookPart.Workbook.Sheets.Cast<Sheet>().ToList();
            if (sheets.Count == 1)
            {
                sheet = sheets.Single();
            }
            else if (sheetName == null)
            {
                var typeName = typeof(T).Name;
                sheet = sheets.Find(s => StringsMatch(s.Name.Value, typeName));
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

            if (!(_document.WorkbookPart.GetPartById(sheet.Id.Value) is WorksheetPart worksheetPart))
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
                var matchingProperties = properties.Where(p => StringsMatch(p.Name, column)).ToList();
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
                                SetItemProperty(item, ((double?)cellValueDoubleObject).Value, propertyName);
                            }

                            break;
                        case "Int32":
                            var cellValueIntObject = GetCellValue<int>(cell, stringTable);
                            if (cellValueIntObject != null)
                            {
                                SetItemProperty(item, ((int?)cellValueIntObject).Value, propertyName);
                            }
                            break;
                        case "Int64":
                            var cellValueLongObject = GetCellValue<long>(cell, stringTable);
                            if (cellValueLongObject != null)
                            {
                                SetItemProperty(item, ((long?)cellValueLongObject).Value, propertyName);
                            }
                            break;
                        case "String":
                            SetItemProperty(item, (string?)GetCellValue<string>(cell, stringTable), propertyName);
                            break;
                        case "Nullable`1<Boolean>":
                            SetItemProperty(item, (bool?)GetCellValue<bool?>(cell, stringTable), propertyName);
                            break;
                        case "Nullable`1<Double>":
                            SetItemProperty(item, (double?)GetCellValue<double?>(cell, stringTable), propertyName);
                            break;
                        case "Nullable`1<Single>":
                            SetItemProperty(item, (float?)GetCellValue<float?>(cell, stringTable), propertyName);
                            break;
                        case "Nullable`1<Int64>":
                            SetItemProperty(item, (long?)GetCellValue<long?>(cell, stringTable), propertyName);
                            break;
                        case "Nullable`1<Int32>":
                            SetItemProperty(item, (int?)GetCellValue<int?>(cell, stringTable), propertyName);
                            break;
                        case "Nullable`1<Int16>":
                            SetItemProperty(item, (short?)GetCellValue<short?>(cell, stringTable), propertyName);
                            break;
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
            while (Numbers.Contains(tweakString[0]))
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
            return ((CellValues)cell.DataType) switch
            {
                CellValues.SharedString => stringTable.SharedStringTable.ElementAt(int.Parse(cellValueText)).InnerText,
                _ => cellValueText ?? cell.InnerText,
            };
        }

        private object? GetCellValueDirect(Cell cell, SharedStringTablePart stringTable)
        {
            var cellValueText = cell.CellValue?.Text;
            if (cell.DataType == null)
            {
                return cellValueText;
            }
            switch ((CellValues)cell.DataType)
            {
                case CellValues.SharedString:
                    return stringTable.SharedStringTable
                        .ElementAt(int.Parse(cellValueText)).InnerText;
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
                    return cellValueText ?? cell.InnerText;
                default:
                    throw new NotSupportedException($"Unsupported data type {cell.DataType?.Value.ToString() ?? "None"}");
            }
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
                        return cellValueText;
                }
            }
            switch ((CellValues)cell.DataType)
            {
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
    }
}