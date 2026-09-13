using DocumentFormat.OpenXml;

namespace PanoramicData.SheetMagic;

/// <summary>
/// Styling and workbook generation methods
/// </summary>
public partial class MagicSpreadsheet
{
	private void GenerateWorkbookStylesPart1Content(WorkbookStylesPart workbookStylesPart1)
	{
		var stylesheet1 = new Stylesheet { MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "x14ac x16r2 xr xr9" } };
		stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
		stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
		stylesheet1.AddNamespaceDeclaration("x16r2", "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main");
		stylesheet1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
		stylesheet1.AddNamespaceDeclaration("xr9", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision9");

		// Adding a new date format
		var numberingFormats = CreateNumberingFormats();
		var dateNumberFormatId = numberingFormats.GetFirstChild<NumberingFormat>()!.NumberFormatId;

		var differentialFormats = new DifferentialFormats { Count = 3U };

		stylesheet1.Append(numberingFormats);
		stylesheet1.Append(CreateDefaultFonts());
		stylesheet1.Append(CreateDefaultFills());
		stylesheet1.Append(CreateDefaultBorders());
		stylesheet1.Append(CreateCellStyleFormats(dateNumberFormatId));
		stylesheet1.Append(CreateCellFormats(dateNumberFormatId));
		stylesheet1.Append(CreateCellStyles());
		stylesheet1.Append(differentialFormats);
		stylesheet1.Append(CreateTableStyles(differentialFormats));
		stylesheet1.Append(CreateDefaultColors());

		workbookStylesPart1.Stylesheet = stylesheet1;
	}

	private static Fonts CreateDefaultFonts()
	{
		var fonts = new Fonts { Count = 1U, KnownFonts = true };
		var font = new Font();
		font.Append(new FontSize { Val = 11D });
		font.Append(new Color { Theme = 1U });
		font.Append(new FontName { Val = "Calibri" });
		font.Append(new FontFamilyNumbering { Val = 2 });
		font.Append(new FontScheme { Val = FontSchemeValues.Minor });
		fonts.Append(font);
		return fonts;
	}

	private static Fills CreateDefaultFills()
	{
		var fills = new Fills { Count = 2U };
		var noneFill = new Fill();
		noneFill.Append(new PatternFill { PatternType = PatternValues.None });
		var gray125Fill = new Fill();
		gray125Fill.Append(new PatternFill { PatternType = PatternValues.Gray125 });
		fills.Append(noneFill);
		fills.Append(gray125Fill);
		return fills;
	}

	private static Borders CreateDefaultBorders()
	{
		var borders = new Borders { Count = 1U };
		var outerBorder = new Border();
		outerBorder.Append(new LeftBorder());
		outerBorder.Append(new RightBorder());
		outerBorder.Append(new TopBorder());
		outerBorder.Append(new BottomBorder());
		outerBorder.Append(new DiagonalBorder());
		borders.Append(outerBorder);
		return borders;
	}

	private static NumberingFormats CreateNumberingFormats()
	{
		var numberingFormats = new NumberingFormats { Count = 1U };
		numberingFormats.Append(new NumberingFormat
		{
			NumberFormatId = 165, // any number greater than 164 will do for custom format
			FormatCode = "yyyy-mm-dd hh:mm:ss"
		});
		return numberingFormats;
	}

	private static CellStyleFormats CreateCellStyleFormats(UInt32Value? dateNumberFormatId)
	{
		var cellStyleFormats = new CellStyleFormats { Count = 1U };
		cellStyleFormats.Append(new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U });
		cellStyleFormats.Append(new CellFormat { NumberFormatId = dateNumberFormatId });
		return cellStyleFormats;
	}

	private static CellFormats CreateCellFormats(UInt32Value? dateNumberFormatId)
	{
		var cellFormats = new CellFormats { Count = 1U };
		cellFormats.Append(new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U });
		cellFormats.Append(new CellFormat { NumberFormatId = dateNumberFormatId });
		return cellFormats;
	}

	private static CellStyles CreateCellStyles()
	{
		var cellStyles = new CellStyles { Count = 1U };
		cellStyles.Append(new CellStyle { Name = "Normal", FormatId = 0U, BuiltinId = 0U });
		return cellStyles;
	}

	private static Colors CreateDefaultColors()
	{
		var colors = new Colors();
		var mruColors = new MruColors();
		mruColors.Append(new Color { Rgb = "FFE1CCF0" });
		colors.Append(mruColors);
		return colors;
	}

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
