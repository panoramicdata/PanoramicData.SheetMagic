﻿using PanoramicData.SheetMagic.Exceptions;
using System.Collections.Generic;

namespace PanoramicData.SheetMagic
{
	public class AddSheetOptions
	{
		public HashSet<string>? IncludeProperties { get; set; }
		public HashSet<string>? ExcludeProperties { get; set; }

		/// <summary>
		/// Whether to sort the combined list of properties, and any additional extended properties. Defaults to true.
		/// </summary>
		public bool SortExtendedProperties { get; set; } = true;

		public TableOptions? TableOptions { get; set; }

		internal void Validate()
		{
			if (IncludeProperties != null && ExcludeProperties != null)
			{
				throw new ValidationException($"Cannot set both {nameof(IncludeProperties)} and {nameof(ExcludeProperties)}");
			}

			TableOptions.Validate();
		}
	}
}