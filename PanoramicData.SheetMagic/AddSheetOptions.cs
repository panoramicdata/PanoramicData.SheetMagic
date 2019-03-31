using PanoramicData.SheetMagic.Exceptions;
using System.Collections.Generic;

namespace PanoramicData.SheetMagic
{
	public class AddSheetOptions
	{
		public HashSet<string> IncludeProperties { get; set; }
		public HashSet<string> ExcludeProperties { get; set; }

		internal void Validate()
		{
			if (IncludeProperties != null && ExcludeProperties != null)
			{
				throw new ValidationException($"Cannot set both {nameof(IncludeProperties)} and {nameof(ExcludeProperties)}");
			}
		}
	}
}