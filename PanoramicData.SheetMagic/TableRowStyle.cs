using System.Drawing;

namespace PanoramicData.SheetMagic
{
	/// <summary>
	/// A table row style
	/// </summary>
	public class TableRowStyle
	{
		/// <summary>
		/// The background color
		/// </summary>
		public Color? BackgroundColor { get; set; }

		/// <summary>
		/// The font color
		/// </summary>
		public Color? FontColor { get; set; }

		/// <summary>
		/// The optional inner border color
		/// </summary>
		public Color? InnerBorderColor { get; set; }

		/// <summary>
		/// The optional outer border color
		/// </summary>
		public Color? OuterBorderColor { get; set; }

		/// <summary>
		/// The font weight
		/// </summary>
		public FontWeight FontWeight { get; set; }
	}
}
