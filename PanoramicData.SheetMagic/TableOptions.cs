namespace PanoramicData.SheetMagic;

/// <summary>
/// Options for configuring Excel table formatting and behavior.
/// </summary>
public class TableOptions
{
	private string _displayName = "Table";
	private string _name = "Table";

	/// <summary>
	/// Gets or sets the internal name of the table. Special characters are replaced with underscores.
	/// </summary>
	public string Name
	{
		get => _name;

		set
		{
			_name = Regex.Replace(value, "[^0-9a-zA-Z]+", "_").Replace(" ", string.Empty);
		}
	}

	/// <summary>
	/// Gets or sets the display name of the table. Special characters are replaced with underscores.
	/// </summary>
	public string DisplayName
	{
		get => _displayName;

		set
		{
			_displayName = Regex.Replace(value, "[^0-9a-zA-Z]+", "_").Replace(" ", string.Empty);
		}
	}

	/// <summary>
	/// Gets or sets the built-in Excel table style to apply.
	/// </summary>
	public XlsxTableStyle XlsxTableStyle { get; set; } = XlsxTableStyle.TableStyleMedium9;

	/// <summary>
	/// Gets or sets whether to show a totals row at the bottom of the table.
	/// </summary>
	public bool ShowTotalsRow { get; set; }

	/// <summary>
	/// Gets or sets whether to apply special formatting to the first column.
	/// </summary>
	public bool ShowFirstColumn { get; set; }

	/// <summary>
	/// Gets or sets whether to apply special formatting to the last column.
	/// </summary>
	public bool ShowLastColumn { get; set; }

	/// <summary>
	/// Gets or sets whether to show alternating row stripes. Defaults to true.
	/// </summary>
	public bool ShowRowStripes { get; set; } = true;

	/// <summary>
	/// Gets or sets whether to show alternating column stripes.
	/// </summary>
	public bool ShowColumnStripes { get; set; }

	/// <summary>
	/// Gets or sets the name of a custom table style to use instead of a built-in style.
	/// </summary>
	public string? CustomTableStyle { get; set; }

	internal void Validate(List<CustomTableStyle> tableStyles)
	{
		if (DisplayName.Contains(' '))
		{
			throw new ValidationException($"TableOptions display name cannot contain spaces. Found '{DisplayName}'.");
		}

		if (CustomTableStyle != null && !tableStyles.Any(ts => ts.Name == CustomTableStyle))
		{
			throw new ValidationException($"Undefined CustomTableStyle '{CustomTableStyle}' was requested. Define it in the {nameof(Options)}.{nameof(Options.TableStyles)} provided in the {nameof(MagicSpreadsheet)} constructor.");
		}
	}
}