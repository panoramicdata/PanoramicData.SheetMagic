using PanoramicData.SheetMagic.Interfaces;

namespace PanoramicData.SheetMagic;

/// <summary>
/// Defines a custom table style with configurable row styles.
/// </summary>
public class CustomTableStyle : IValidate
{
	/// <summary>
	/// Gets or sets the name of the custom table style.
	/// </summary>
	public string Name { get; set; } = "Custom Table Style";

	/// <summary>
	/// Gets or sets the style for the header row.
	/// </summary>
	public TableRowStyle? HeaderRowStyle { get; set; }

	/// <summary>
	/// Gets or sets the style for odd-numbered data rows.
	/// </summary>
	public TableRowStyle? OddRowStyle { get; set; }

	/// <summary>
	/// Gets or sets the style for even-numbered data rows.
	/// </summary>
	public TableRowStyle? EvenRowStyle { get; set; }

	/// <summary>
	/// Gets or sets the style applied to the entire table.
	/// </summary>
	public TableRowStyle? WholeTableStyle { get; set; }

	/// <summary>
	/// Validates the custom table style configuration.
	/// </summary>
	/// <exception cref="ValidationException">Thrown when the style has no name or no styles defined.</exception>
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
