using PanoramicData.SheetMagic.Exceptions;
using System.Collections.Generic;

namespace PanoramicData.SheetMagic
{
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

		public void Validate(List<CustomTableStyle> tableStyles)
		{
			if (IncludeProperties != null && ExcludeProperties != null)
			{
				throw new ValidationException($"Cannot set both {nameof(IncludeProperties)} and {nameof(ExcludeProperties)}");
			}

			TableOptions?.Validate(tableStyles);
		}
	}
}