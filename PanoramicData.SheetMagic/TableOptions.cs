using PanoramicData.SheetMagic.Exceptions;

namespace PanoramicData.SheetMagic
{
	public class TableOptions
	{
		public string Name { get; set; } = "Table";
		public string DisplayName { get; set; } = "Table";
		public XlsxTableStyle XlsxTableStyle { get; set; } = XlsxTableStyle.TableStyleMedium9;
		public bool ShowTotalsRow { get; set; }
		public bool ShowFirstColumn { get; set; }
		public bool ShowLastColumn { get; set; }
		public bool ShowRowStripes { get; set; } = true;
		public bool ShowColumnStripes { get; set; }

		internal void Validate()
		{
			if (DisplayName.Contains(" "))
			{
				throw new ValidationException($"TableOptions display name cannot contain spaces.  Found '{DisplayName}'.");
			}
		}
	}
}