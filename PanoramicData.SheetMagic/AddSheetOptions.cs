using PanoramicData.SheetMagic.Exceptions;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace PanoramicData.SheetMagic;

public class AddSheetOptions
{
	/// <summary>
	/// The properties to include
	/// </summary>
	public HashSet<string>? IncludeProperties { get; set; }

	/// <summary>
	/// The properties to exclude
	/// </summary>
	public HashSet<string>? ExcludeProperties { get; set; }

	/// <summary>
	/// Whether to sort the combined list of properties, and any additional extended properties. Defaults to true.
	/// </summary>
	public bool SortExtendedProperties { get; set; } = true;

	/// <summary>
	/// TableOptions
	/// </summary>
	public TableOptions? TableOptions { get; set; } = new TableOptions
	{
		XlsxTableStyle = XlsxTableStyle.TableStyleMedium11
	};

	/// <summary>
	/// An optional EnumerableCellOptions.  If not set, the Options EnumerableCellOptions set in Options is used.
	/// </summary>
	public EnumerableCellOptions? EnumerableCellOptions { get; set; }

	/// <summary>
	/// In Excel, it is not possible to add a table with no rows.
	/// If the user tries to add a table with no rows and this property is set to:
	/// - true (default): SheetMagic will throw an InvalidOperationException if
	/// - false: SheetMagic will silently not add a new sheet
	/// </summary>
	public bool ThrowExceptionOnEmptyList { get; set; } = true;

	public void Validate(List<CustomTableStyle> tableStyles)
	{
		if (IncludeProperties != null && ExcludeProperties != null)
		{
			throw new ValidationException($"Cannot set both {nameof(IncludeProperties)} and {nameof(ExcludeProperties)}");
		}

		TableOptions?.Validate(tableStyles);
	}

	internal AddSheetOptions Clone()
		=> new()
		{
			EnumerableCellOptions = EnumerableCellOptions == null
				? null
				: new EnumerableCellOptions
				{
					CellDelimiter = EnumerableCellOptions.CellDelimiter,
					Expand = EnumerableCellOptions.Expand,
				},
			ExcludeProperties = ExcludeProperties == null
				? null
				: new HashSet<string>(ExcludeProperties),
			IncludeProperties = IncludeProperties == null
				? null
				: new HashSet<string>(IncludeProperties),
			SortExtendedProperties = SortExtendedProperties,
			TableOptions = TableOptions == null
				? null
				: new TableOptions
				{
					CustomTableStyle = TableOptions.CustomTableStyle,
					DisplayName = TableOptions.DisplayName,
					Name = TableOptions.Name,
					ShowColumnStripes = TableOptions.ShowColumnStripes,
					ShowFirstColumn = TableOptions.ShowFirstColumn,
					ShowLastColumn = TableOptions.ShowLastColumn,
					ShowRowStripes = TableOptions.ShowRowStripes,
					ShowTotalsRow = TableOptions.ShowTotalsRow,
					XlsxTableStyle = TableOptions.XlsxTableStyle
				},
			ThrowExceptionOnEmptyList = ThrowExceptionOnEmptyList
		};
}