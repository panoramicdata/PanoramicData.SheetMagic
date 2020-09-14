using PanoramicData.SheetMagic.Exceptions;
using PanoramicData.SheetMagic.Interfaces;

namespace PanoramicData.SheetMagic
{
	public class CustomTableStyle : IValidate
	{
		public string Name { get; set; } = "Custom Table Style";

		public TableRowStyle? HeaderRowStyle { get; set; }

		public TableRowStyle? OddRowStyle { get; set; }

		public TableRowStyle? EvenRowStyle { get; set; }

		public TableRowStyle? WholeTableStyle { get; set; }

		public void Validate()
		{
			if (string.IsNullOrWhiteSpace(Name))
			{
				throw new ValidationException("CustomTableStyle with no name is present.");
			}

			if (HeaderRowStyle is null
				&& OddRowStyle is null
				&& EvenRowStyle is null
				&& WholeTableStyle is null)
			{
				throw new ValidationException($"No style set in CustomTableStyle '{Name}'.");
			}
		}
	}
}
