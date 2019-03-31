using System.Collections.Generic;

namespace PanoramicData.SheetMagic
{
	public class Extended<T>
	{
		public T Item { get; set; }
		public Dictionary<string, object> Properties { get; set; }
	}
}