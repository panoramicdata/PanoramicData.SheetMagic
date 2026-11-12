using DocumentFormat.OpenXml;

namespace PanoramicData.SheetMagic;

/// <summary>
/// Table style and colour helpers
/// </summary>
public partial class MagicSpreadsheet
{
	/// <summary>
	/// Creates the workbook's table styles, adding a differential format to
	/// <paramref name="differentialFormats"/> for each row style of the first custom table style.
	/// </summary>
	private TableStyles CreateTableStyles(DifferentialFormats differentialFormats)
	{
		var tableStyles = new TableStyles { Count = 1U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

		if (_options.TableStyles.Count == 0)
		{
			return tableStyles;
		}

		var customTableStyle = _options.TableStyles[0];

		var tableStyle = new TableStyle
		{
			Name = customTableStyle.Name,
			Pivot = false,
			Count = CountRowStyles(customTableStyle)
		};
		tableStyle.SetAttribute(new OpenXmlAttribute("xr9", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision9", "{640A183E-9F4E-4A71-80D9-2176963C18AB}"));
		tableStyles.Append(tableStyle);

		var tableStyleIndex = 0U;
		AddTableStyleElement(customTableStyle.OddRowStyle, differentialFormats, tableStyle, tableStyleIndex++, TableStyleValues.FirstRowStripe);
		AddTableStyleElement(customTableStyle.EvenRowStyle, differentialFormats, tableStyle, tableStyleIndex++, TableStyleValues.SecondRowStripe);
		AddTableStyleElement(customTableStyle.HeaderRowStyle, differentialFormats, tableStyle, tableStyleIndex++, TableStyleValues.HeaderRow);
		AddTableStyleElement(customTableStyle.WholeTableStyle, differentialFormats, tableStyle, tableStyleIndex, TableStyleValues.WholeTable);

		return tableStyles;
	}

	private static uint CountRowStyles(CustomTableStyle customTableStyle)
		=> (uint)new TableRowStyle?[]
		{
			customTableStyle.OddRowStyle,
			customTableStyle.EvenRowStyle,
			customTableStyle.HeaderRowStyle,
			customTableStyle.WholeTableStyle
		}.Count(static style => style is not null);

	private static void AddTableStyleElement(
		TableRowStyle? thisCustomTableStyle,
		DifferentialFormats differentialFormats,
		TableStyle tableStyle1,
		uint tableStyleIndex,
		TableStyleValues tableStyleValues)
	{
		if (thisCustomTableStyle is null)
		{
			return;
		}

		var differentialFormat = new DifferentialFormat();

		if (thisCustomTableStyle.FontColor.HasValue)
		{
			differentialFormat.Append(CreateTableStyleFont(thisCustomTableStyle));
		}

		if (thisCustomTableStyle.BackgroundColor.HasValue)
		{
			differentialFormat.Append(CreateTableStyleFill(thisCustomTableStyle.BackgroundColor.Value));
		}

		var border = CreateTableStyleBorder(thisCustomTableStyle);
		if (border is not null)
		{
			differentialFormat.Append(border);
		}

		differentialFormats.Append(differentialFormat);
		tableStyle1.Append(new TableStyleElement { Type = tableStyleValues, FormatId = tableStyleIndex });
	}

	private static Font CreateTableStyleFont(TableRowStyle tableRowStyle)
	{
		var font = new Font();
		if (tableRowStyle.FontWeight == FontWeight.Bold)
		{
			font.Append(new Bold());
		}

		font.Append(GetColor(tableRowStyle.FontColor!.Value));
		return font;
	}

	private static Fill CreateTableStyleFill(System.Drawing.Color backgroundColor)
	{
		var fill = new Fill();
		var patternFill = new PatternFill();
		patternFill.Append(new BackgroundColor { Rgb = GetHexBinaryValue(backgroundColor) });
		fill.Append(patternFill);
		return fill;
	}

	/// <summary>
	/// Creates the border for a table row style, or null where the style defines no border.
	/// </summary>
	private static Border? CreateTableStyleBorder(TableRowStyle tableRowStyle)
	{
		if (!tableRowStyle.InnerBorderColor.HasValue && !tableRowStyle.OuterBorderColor.HasValue)
		{
			return null;
		}

		var border = new Border();

		if (tableRowStyle.OuterBorderColor.HasValue)
		{
			var outerColor = tableRowStyle.OuterBorderColor.Value;
			border.Append(new LeftBorder { Color = GetColor(outerColor), Style = BorderStyleValues.Thin });
			border.Append(new RightBorder { Color = GetColor(outerColor), Style = BorderStyleValues.Thin });
			border.Append(new TopBorder { Color = GetColor(outerColor), Style = BorderStyleValues.Thin });
			border.Append(new BottomBorder { Color = GetColor(outerColor), Style = BorderStyleValues.Thin });
		}

		if (tableRowStyle.InnerBorderColor.HasValue)
		{
			var innerColor = tableRowStyle.InnerBorderColor.Value;
			border.Append(new VerticalBorder { Color = GetColor(innerColor), Style = BorderStyleValues.Thin });
			border.Append(new HorizontalBorder { Color = GetColor(innerColor), Style = BorderStyleValues.Thin });
		}

		return border;
	}

	private static Color GetColor(System.Drawing.Color color)
		=> Equals(color, System.Drawing.Color.White)
			? new Color { Theme = 0U }
			: new Color { Rgb = GetHexBinaryValue(color) };

	private static HexBinaryValue GetHexBinaryValue(System.Drawing.Color color) => new()
	{
		Value = $"FF{color.R:X2}{color.G:X2}{color.B:X2}"
	};
}
