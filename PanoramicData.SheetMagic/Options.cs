using PanoramicData.SheetMagic.Interfaces;
using System.Collections.Generic;

namespace PanoramicData.SheetMagic;

public class Options : IValidate
{
	/// <summary>
	/// Whether to stop processing on the first empty row in the table
	/// </summary>
	public bool StopProcessingOnFirstEmptyRow { get; set; }

	/// <summary>
	/// Whether to interpret an empty row as null
	/// </summary>
	public bool EmptyRowInterpretedAsNull { get; set; }

	/// <summary>
	/// Whether to ignore unmapped properties
	/// </summary>
	public bool IgnoreUnmappedProperties { get; set; }

	/// <summary>
	/// Whether to load null extended properties
	/// </summary>
	public bool LoadNullExtendedProperties { get; set; }

	/// <summary>
	/// When using AddSheet, if no addSheetOptions are specified, the default AddSheetOptions to use.
	/// Defaults to a reasonable set of options.
	/// </summary>
	public AddSheetOptions DefaultAddSheetOptions { get; set; } = new AddSheetOptions();

	/// <summary>
	/// Custom table styles
	/// </summary>
	public List<CustomTableStyle> TableStyles { get; set; } = [];

	/// <summary>
	/// EnumerableCellOptions
	/// </summary>
	public EnumerableCellOptions EnumerableCellOptions { get; set; } = new EnumerableCellOptions();

	public string ListSeparator { get; set; } = ", ";

	public void Validate()
	{
		foreach (var tableStyle in TableStyles)
		{
			tableStyle.Validate();
		}
	}
}