using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

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

		// Fonts
		var fonts = new Fonts { Count = 1U, KnownFonts = true };
		var font = new Font();
		font.Append(new FontSize { Val = 11D });
		font.Append(new Color { Theme = 1U });
		font.Append(new FontName { Val = "Calibri" });
		font.Append(new FontFamilyNumbering { Val = 2 });
		font.Append(new FontScheme { Val = FontSchemeValues.Minor });
		fonts.Append(font);

		// Fills
		var fills = new Fills { Count = 2U };
		var noneFill = new Fill();
		noneFill.Append(new PatternFill { PatternType = PatternValues.None });
		var gray125Fill = new Fill();
		gray125Fill.Append(new PatternFill { PatternType = PatternValues.Gray125 });
		fills.Append(noneFill);
		fills.Append(gray125Fill);

		// Outer Borders
		var borders = new Borders { Count = 1U };
		var outerBorder = new Border();
		outerBorder.Append(new LeftBorder());
		outerBorder.Append(new RightBorder());
		outerBorder.Append(new TopBorder());
		outerBorder.Append(new BottomBorder());
		outerBorder.Append(new DiagonalBorder());
		borders.Append(outerBorder);

		// Adding a new date format
		var nf = new NumberingFormat
		{
			NumberFormatId = 165, // any number greater than 164 will do for custom format
			FormatCode = "yyyy-mm-dd hh:mm:ss"
		};
		var numberingFormats = new NumberingFormats { Count = 1U };
		numberingFormats.Append(nf);

		var cellStyleFormats1 = new CellStyleFormats { Count = 1U };
		var cellFormat1 = new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U };
		var csf = new CellFormat
		{
			NumberFormatId = nf.NumberFormatId
		};
		cellStyleFormats1.Append(cellFormat1);
		cellStyleFormats1.Append(csf);

		var cellFormats1 = new CellFormats { Count = 1U };
		var cellFormat2 = new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U };
		var cf = new CellFormat
		{
			NumberFormatId = csf.NumberFormatId
		};

		cellFormats1.Append(cellFormat2);
		cellFormats1.Append(cf);

		var cellStyles1 = new CellStyles { Count = 1U };
		var cellStyle1 = new CellStyle { Name = "Normal", FormatId = 0U, BuiltinId = 0U };

		cellStyles1.Append(cellStyle1);

		var differentialFormats = new DifferentialFormats { Count = 3U };
		var tableStyles1 = new TableStyles { Count = 1U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

		if (_options.TableStyles.Count > 0)
		{
			CustomTableStyle customTableStyle = _options.TableStyles[0];

			var tableStyleCount = 0U;
			if (customTableStyle.OddRowStyle != null)
			{
				tableStyleCount++;
			}

			if (customTableStyle.EvenRowStyle != null)
			{
				tableStyleCount++;
			}

			if (customTableStyle.HeaderRowStyle != null)
			{
				tableStyleCount++;
			}

			if (customTableStyle.WholeTableStyle != null)
			{
				tableStyleCount++;
			}

			var tableStyle1 = new TableStyle { Name = customTableStyle.Name, Pivot = false, Count = tableStyleCount };
			tableStyle1.SetAttribute(new OpenXmlAttribute("xr9", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision9", "{640A183E-9F4E-4A71-80D9-2176963C18AB}"));
			tableStyles1.Append(tableStyle1);
			var tableStyleIndex = 0U;
			AddTableStyleElement(customTableStyle.OddRowStyle, differentialFormats, tableStyle1, tableStyleIndex++, TableStyleValues.FirstRowStripe);
			AddTableStyleElement(customTableStyle.EvenRowStyle, differentialFormats, tableStyle1, tableStyleIndex++, TableStyleValues.SecondRowStripe);
			AddTableStyleElement(customTableStyle.HeaderRowStyle, differentialFormats, tableStyle1, tableStyleIndex++, TableStyleValues.HeaderRow);
			AddTableStyleElement(customTableStyle.WholeTableStyle, differentialFormats, tableStyle1, tableStyleIndex, TableStyleValues.WholeTable);
		}
		// Colors
		var colors1 = new Colors();

		var mruColors1 = new MruColors();
		var color5 = new Color { Rgb = "FFE1CCF0" };

		mruColors1.Append(color5);

		colors1.Append(mruColors1);

		stylesheet1.Append(numberingFormats);
		stylesheet1.Append(fonts);
		stylesheet1.Append(fills);
		stylesheet1.Append(borders);
		stylesheet1.Append(cellStyleFormats1);
		stylesheet1.Append(cellFormats1);
		stylesheet1.Append(cellStyles1);
		stylesheet1.Append(differentialFormats);
		stylesheet1.Append(tableStyles1);
		stylesheet1.Append(colors1);

		workbookStylesPart1.Stylesheet = stylesheet1;
	}

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

		// Font color
		if (thisCustomTableStyle.FontColor.HasValue)
		{
			var font = new Font();
			if (thisCustomTableStyle.FontWeight == FontWeight.Bold)
			{
				font.Append(new Bold());
			}

			font.Append(GetColor(thisCustomTableStyle.FontColor.Value));
			differentialFormat.Append(font);
		}

		// Background color
		if (thisCustomTableStyle.BackgroundColor.HasValue)
		{
			var fill = new Fill();
			var patternFill = new PatternFill();
			patternFill.Append(new BackgroundColor { Rgb = GetHexBinaryValue(thisCustomTableStyle.BackgroundColor.Value) });
			fill.Append(patternFill);
			differentialFormat.Append(fill);
		}

		// Inner border
		if (thisCustomTableStyle.InnerBorderColor.HasValue || thisCustomTableStyle.OuterBorderColor.HasValue)
		{
			var border = new Border();

			if (thisCustomTableStyle.OuterBorderColor.HasValue)
			{
				border.Append(new LeftBorder { Color = GetColor(thisCustomTableStyle.OuterBorderColor.Value), Style = BorderStyleValues.Thin });
				border.Append(new RightBorder { Color = GetColor(thisCustomTableStyle.OuterBorderColor.Value), Style = BorderStyleValues.Thin });
				border.Append(new TopBorder { Color = GetColor(thisCustomTableStyle.OuterBorderColor.Value), Style = BorderStyleValues.Thin });
				border.Append(new BottomBorder { Color = GetColor(thisCustomTableStyle.OuterBorderColor.Value), Style = BorderStyleValues.Thin });
			}

			if (thisCustomTableStyle.InnerBorderColor.HasValue)
			{
				border.Append(new VerticalBorder { Color = GetColor(thisCustomTableStyle.InnerBorderColor.Value), Style = BorderStyleValues.Thin });
				border.Append(new HorizontalBorder { Color = GetColor(thisCustomTableStyle.InnerBorderColor.Value), Style = BorderStyleValues.Thin });
			}

			differentialFormat ??= new DifferentialFormat();
			differentialFormat.Append(border);
		}

		differentialFormats.Append(differentialFormat);
		tableStyle1.Append(new TableStyleElement { Type = tableStyleValues, FormatId = tableStyleIndex });
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
