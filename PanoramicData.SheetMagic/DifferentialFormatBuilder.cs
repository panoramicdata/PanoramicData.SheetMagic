namespace PanoramicData.SheetMagic;

/// <summary>
/// Builds the differential formats applied by conditional formatting rules
/// </summary>
internal static class DifferentialFormatBuilder
{
	internal static DifferentialFormat CreateDifferentialFormatFromStyle(ConditionalFormatStyle style, ref uint nextNumberFormatId)
	{
		var dxf = new DifferentialFormat();

		var font = CreateDifferentialFont(style);
		if (font is not null)
		{
			dxf.Append(font);
		}

		if (style.BackgroundColor.HasValue)
		{
			dxf.Append(CreateSolidFill(style.BackgroundColor.Value));
		}

		if (style.BorderColor.HasValue)
		{
			dxf.Append(CreateThinBorder(style.BorderColor.Value));
		}

		if (style.NumberFormat is not null)
		{
			dxf.Append(new NumberingFormat
			{
				NumberFormatId = nextNumberFormatId++,
				FormatCode = style.NumberFormat
			});
		}

		return dxf;
	}

	internal static Font? CreateDifferentialFont(ConditionalFormatStyle style)
	{
		if (!HasFontStyling(style))
		{
			return null;
		}

		var font = new Font();
		if (style.FontWeight == FontWeight.Bold)
		{
			font.Append(new Bold());
		}

		if (style.Italic == true)
		{
			font.Append(new Italic());
		}

		if (style.Strikethrough == true)
		{
			font.Append(new Strike());
		}

		if (style.FontColor.HasValue)
		{
			font.Append(MagicSpreadsheet.GetColor(style.FontColor.Value));
		}

		return font;
	}

	internal static bool HasFontStyling(ConditionalFormatStyle style)
		=> style.FontColor.HasValue
			|| style.FontWeight.HasValue
			|| style.Italic.HasValue
			|| style.Strikethrough.HasValue;

	internal static Fill CreateSolidFill(System.Drawing.Color color)
	{
		var fill = new Fill();
		var patternFill = new PatternFill { PatternType = PatternValues.Solid };
		patternFill.Append(new ForegroundColor { Rgb = MagicSpreadsheet.GetHexBinaryValue(color) });
		fill.Append(patternFill);
		return fill;
	}

	internal static Border CreateThinBorder(System.Drawing.Color color)
	{
		var border = new Border();
		border.Append(new LeftBorder { Color = MagicSpreadsheet.GetColor(color), Style = BorderStyleValues.Thin });
		border.Append(new RightBorder { Color = MagicSpreadsheet.GetColor(color), Style = BorderStyleValues.Thin });
		border.Append(new TopBorder { Color = MagicSpreadsheet.GetColor(color), Style = BorderStyleValues.Thin });
		border.Append(new BottomBorder { Color = MagicSpreadsheet.GetColor(color), Style = BorderStyleValues.Thin });
		return border;
	}
}
