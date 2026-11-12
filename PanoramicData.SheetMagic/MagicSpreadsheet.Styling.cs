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
}
