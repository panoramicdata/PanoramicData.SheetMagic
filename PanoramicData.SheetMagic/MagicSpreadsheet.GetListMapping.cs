using PanoramicData.SheetMagic.Extensions;

namespace PanoramicData.SheetMagic;

/// <summary>
/// Reading sheets - column mapping and row processing
/// </summary>
public partial class MagicSpreadsheet
{
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

			if (ShouldSkipEmptyRow(cells, stringTable))
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

	private bool ShouldSkipEmptyRow(List<Cell> cells, SharedStringTablePart? stringTable)
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
