namespace PanoramicData.SheetMagic
{
	public class Options
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

		public bool LoadNullExtendedProperties { get; internal set; }
	}
}