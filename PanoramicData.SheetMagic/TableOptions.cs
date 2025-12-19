namespace PanoramicData.SheetMagic;

public class TableOptions
{
	private string _displayName = "Table";
	private string _name = "Table";

	public string Name
	{
		get => _name;

		set
		{
			_name = Regex.Replace(value, "[^0-9a-zA-Z]+", "_").Replace(" ", string.Empty);
		}
	}

	public string DisplayName
	{
		get => _displayName;

		set
		{
			_displayName = Regex.Replace(value, "[^0-9a-zA-Z]+", "_").Replace(" ", string.Empty);
		}
	}

	public XlsxTableStyle XlsxTableStyle { get; set; } = XlsxTableStyle.TableStyleMedium9;

	public bool ShowTotalsRow { get; set; }

	public bool ShowFirstColumn { get; set; }

	public bool ShowLastColumn { get; set; }

	public bool ShowRowStripes { get; set; } = true;

	public bool ShowColumnStripes { get; set; }

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