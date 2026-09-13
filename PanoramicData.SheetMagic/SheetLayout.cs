namespace PanoramicData.SheetMagic;

/// <summary>
/// The column layout that was written for a sheet: what each column holds, and how many there are.
/// </summary>
/// <param name="PropertyList">The properties written, in column order, before any extended properties.</param>
/// <param name="ColumnConfigurations">The sheet's column configurations, one per column.</param>
/// <param name="KeyList">The extended property keys written, in column order, after the properties.</param>
/// <param name="TotalColumnCount">The total number of columns written.</param>
internal sealed record SheetLayout(
	List<PropertyInfo> PropertyList,
	Columns ColumnConfigurations,
	List<string> KeyList,
	uint TotalColumnCount);
