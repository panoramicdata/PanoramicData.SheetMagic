namespace PanoramicData.SheetMagic
{
	/// <summary>
	/// The behaviour for cells with IEnumerable properties
	/// </summary>
	public class EnumerableCellOptions
	{
		/// <summary>
		/// Whether to expand IEnumerables, overriding the value set in the Options.
		/// A cell for an IEnumerable property will contain:
		/// true: a delimited version of each ToString of the enumerable, using the CellDelimiter property as a the delimiter;
		/// false: the ToString() version of the enumerable or "NULL" if null.
		/// </summary>
		public bool Expand { get; set; } = true;

		/// <summary>
		/// The delimiter to use in cells, overriding the value set in the Options.
		/// See ExpandIEnumerableCells.
		/// </summary>
		public string? CellDelimiter { get; set; } = ", ";
	}
}