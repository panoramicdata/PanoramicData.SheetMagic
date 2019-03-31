using DocumentFormat.OpenXml;

namespace PanoramicData.SheetMagic
{
	public class TableOptions
	{
		public string Name { get; set; } = "Table";
		public string DisplayName { get; set; } = "Table";
		public XlsxTableStyle XlsxTableStyle { get; set; } = XlsxTableStyle.TableStyleMedium9;
		public BooleanValue TotalsRowShown { get; set; } = false;
		public BooleanValue ShowFirstColumn { get; set; } = false;
		public BooleanValue ShowLastColumn { get; set; } = false;
		public BooleanValue ShowRowStripes { get; set; } = true;
		public BooleanValue ShowColumnStripes { get; set; } = false;
	}
}