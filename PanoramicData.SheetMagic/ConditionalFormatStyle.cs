namespace PanoramicData.SheetMagic;

/// <summary>
/// Defines the formatting style to apply when a conditional format rule is satisfied.
/// Maps to an Excel DifferentialFormat (dxf).
/// </summary>
/// <example>
/// <code>
/// var style = new ConditionalFormatStyle
/// {
///     FontColor = System.Drawing.Color.Green,
///     FontWeight = FontWeight.Bold,
///     BackgroundColor = System.Drawing.Color.LightYellow
/// };
/// </code>
/// </example>
public class ConditionalFormatStyle
{
	/// <summary>
	/// Font color to apply.
	/// </summary>
	public System.Drawing.Color? FontColor { get; set; }

	/// <summary>
	/// Font weight (Bold/Normal) to apply.
	/// </summary>
	public FontWeight? FontWeight { get; set; }

	/// <summary>
	/// Whether to apply italic formatting.
	/// </summary>
	public bool? Italic { get; set; }

	/// <summary>
	/// Whether to apply strikethrough formatting.
	/// </summary>
	public bool? Strikethrough { get; set; }

	/// <summary>
	/// Background fill color to apply.
	/// </summary>
	public System.Drawing.Color? BackgroundColor { get; set; }

	/// <summary>
	/// Border color to apply (all four sides).
	/// </summary>
	public System.Drawing.Color? BorderColor { get; set; }

	/// <summary>
	/// Number format string to apply (e.g. "0.00", "#,##0").
	/// </summary>
	public string? NumberFormat { get; set; }

	internal bool HasFormatting()
		=> FontColor.HasValue ||
			FontWeight == PanoramicData.SheetMagic.FontWeight.Bold ||
			Italic == true ||
			Strikethrough == true ||
			BackgroundColor.HasValue ||
			BorderColor.HasValue ||
			!string.IsNullOrWhiteSpace(NumberFormat);
}
