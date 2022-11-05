using PanoramicData.SheetMagic.Exceptions;
using System.Collections.Generic;
using System.Linq;

namespace PanoramicData.SheetMagic;

public class TableOptions
{
	public string Name { get; set; } = "Table";

	public string DisplayName { get; set; } = "Table";

	public XlsxTableStyle XlsxTableStyle { get; set; } = XlsxTableStyle.TableStyleMedium9;

	public bool ShowTotalsRow { get; set; }

	public bool ShowFirstColumn { get; set; }

	public bool ShowLastColumn { get; set; }

	public bool ShowRowStripes { get; set; } = true;

	public bool ShowColumnStripes { get; set; }

	public string? CustomTableStyle { get; set; } = null;

	internal void Validate(List<CustomTableStyle> tableStyles)
	{
		if (DisplayName.Contains(" "))
		{
			throw new ValidationException($"TableOptions display name cannot contain spaces. Found '{DisplayName}'.");
		}

		if (CustomTableStyle != null && !tableStyles.Any(ts => ts.Name == CustomTableStyle))
		{
			throw new ValidationException($"Undefined CustomTableStyle '{CustomTableStyle}' was requested. Define it in the {nameof(Options)}.{nameof(Options.TableStyles)} provided in the {nameof(MagicSpreadsheet)} constructor.");
		}
	}
}