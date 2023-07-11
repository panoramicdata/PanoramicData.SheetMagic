using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
using X15ac = DocumentFormat.OpenXml.Office2013.ExcelAc;

namespace GeneratedCode;

static internal class SimplestTableFactory
{
	// Creates a SpreadsheetDocument.
	public static void CreatePackage(string filePath)
	{
		using var package = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
		CreateParts(package);
	}

	// Adds child parts and generates content of the specified part.
	private static void CreateParts(SpreadsheetDocument document)
	{
		var extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
		GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

		var workbookPart1 = document.AddWorkbookPart();
		GenerateWorkbookPart1Content(workbookPart1);

		var workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId3");
		GenerateWorkbookStylesPart1Content(workbookStylesPart1);

		var themePart1 = workbookPart1.AddNewPart<ThemePart>("rId2");
		GenerateThemePart1Content(themePart1);

		var worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
		GenerateWorksheetPart1Content(worksheetPart1);

		var tableDefinitionPart1 = worksheetPart1.AddNewPart<TableDefinitionPart>("rId1");
		GenerateTableDefinitionPart1Content(tableDefinitionPart1);

		var sharedStringTablePart1 = workbookPart1.AddNewPart<SharedStringTablePart>("rId4");
		GenerateSharedStringTablePart1Content(sharedStringTablePart1);

		SetPackageProperties(document);
	}

	// Generates content of extendedFilePropertiesPart1.
	private static void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
	{
		var properties1 = new Ap.Properties();
		properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
		var application1 = new Ap.Application
		{
			Text = "Microsoft Excel"
		};
		var documentSecurity1 = new Ap.DocumentSecurity
		{
			Text = "0"
		};
		var scaleCrop1 = new Ap.ScaleCrop
		{
			Text = "false"
		};

		var headingPairs1 = new Ap.HeadingPairs();

		var vTVector1 = new Vt.VTVector
		{
			BaseType = Vt.VectorBaseValues.Variant,
			Size = (UInt32Value)2U
		};

		var variant1 = new Vt.Variant();
		var vTLPSTR1 = new Vt.VTLPSTR
		{
			Text = "Worksheets"
		};

		variant1.Append(vTLPSTR1);

		var variant2 = new Vt.Variant();
		var vTInt321 = new Vt.VTInt32
		{
			Text = "1"
		};

		variant2.Append(vTInt321);

		vTVector1.Append(variant1);
		vTVector1.Append(variant2);

		headingPairs1.Append(vTVector1);

		var titlesOfParts1 = new Ap.TitlesOfParts();

		var vTVector2 = new Vt.VTVector
		{
			BaseType = Vt.VectorBaseValues.Lpstr,
			Size = (UInt32Value)1U
		};
		var vTLPSTR2 = new Vt.VTLPSTR
		{
			Text = "Sheet1"
		};

		vTVector2.Append(vTLPSTR2);

		titlesOfParts1.Append(vTVector2);
		var company1 = new Ap.Company
		{
			Text = ""
		};
		var linksUpToDate1 = new Ap.LinksUpToDate
		{
			Text = "false"
		};
		var sharedDocument1 = new Ap.SharedDocument
		{
			Text = "false"
		};
		var hyperlinksChanged1 = new Ap.HyperlinksChanged
		{
			Text = "false"
		};
		var applicationVersion1 = new Ap.ApplicationVersion
		{
			Text = "16.0300"
		};

		properties1.Append(application1);
		properties1.Append(documentSecurity1);
		properties1.Append(scaleCrop1);
		properties1.Append(headingPairs1);
		properties1.Append(titlesOfParts1);
		properties1.Append(company1);
		properties1.Append(linksUpToDate1);
		properties1.Append(sharedDocument1);
		properties1.Append(hyperlinksChanged1);
		properties1.Append(applicationVersion1);

		extendedFilePropertiesPart1.Properties = properties1;
	}

	// Generates content of workbookPart1.
	private static void GenerateWorkbookPart1Content(WorkbookPart workbookPart1)
	{
		var workbook1 = new Workbook
		{
			MCAttributes = new MarkupCompatibilityAttributes
			{
				Ignorable = "x15 xr xr6 xr10 xr2"
			}
		};
		workbook1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
		workbook1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
		workbook1.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
		workbook1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
		workbook1.AddNamespaceDeclaration("xr6", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6");
		workbook1.AddNamespaceDeclaration("xr10", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10");
		workbook1.AddNamespaceDeclaration("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");
		var fileVersion1 = new FileVersion
		{
			ApplicationName = "xl",
			LastEdited = "7",
			LowestEdited = "7",
			BuildVersion = "22730"
		};
		var workbookProperties1 = new WorkbookProperties
		{
			DefaultThemeVersion = (UInt32Value)166925U
		};

		var alternateContent1 = new AlternateContent();
		alternateContent1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

		var alternateContentChoice1 = new AlternateContentChoice { Requires = "x15" };

		var absolutePath1 = new X15ac.AbsolutePath { Url = "C:\\Users\\david.bond.000\\Downloads\\" };
		absolutePath1.AddNamespaceDeclaration("x15ac", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac");

		alternateContentChoice1.Append(absolutePath1);

		alternateContent1.Append(alternateContentChoice1);

		var openXmlUnknownElement1 = workbookPart1.CreateUnknownElement("<xr:revisionPtr revIDLastSave=\"0\" documentId=\"8_{D93C4835-9CB1-4F2E-BEBC-8B38B516D332}\" xr6:coauthVersionLast=\"45\" xr6:coauthVersionMax=\"45\" xr10:uidLastSave=\"{00000000-0000-0000-0000-000000000000}\" xmlns:xr10=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision10\" xmlns:xr6=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision6\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" />");

		var bookViews1 = new BookViews();

		var workbookView1 = new WorkbookView
		{
			XWindow = -105,
			YWindow = -17880,
			WindowWidth = (UInt32Value)28800U,
			WindowHeight = (UInt32Value)15885U
		};
		workbookView1.SetAttribute(new OpenXmlAttribute("xr2", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2", "{7979DC9F-D70F-4F5A-B6B4-9CBC88F81F24}"));

		bookViews1.Append(workbookView1);

		var sheets1 = new Sheets();
		var sheet1 = new Sheet
		{
			Name = "Sheet1",
			SheetId = (UInt32Value)1U,
			Id = "rId1"
		};

		sheets1.Append(sheet1);
		var calculationProperties1 = new CalculationProperties { CalculationId = (UInt32Value)181029U };

		var workbookExtensionList1 = new WorkbookExtensionList();

		var workbookExtension1 = new WorkbookExtension
		{
			Uri = "{140A7094-0E35-4892-8432-C4D2E57EDEB5}"
		};
		workbookExtension1.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
		var workbookProperties2 = new X15.WorkbookProperties { ChartTrackingReferenceBase = true };

		workbookExtension1.Append(workbookProperties2);

		var workbookExtension2 = new WorkbookExtension
		{
			Uri = "{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}"
		};
		workbookExtension2.AddNamespaceDeclaration("xcalcf", "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures");

		var openXmlUnknownElement2 = workbookPart1.CreateUnknownElement("<xcalcf:calcFeatures xmlns:xcalcf=\"http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures\"><xcalcf:feature name=\"microsoft.com:RD\" /><xcalcf:feature name=\"microsoft.com:FV\" /></xcalcf:calcFeatures>");

		workbookExtension2.Append(openXmlUnknownElement2);

		workbookExtensionList1.Append(workbookExtension1);
		workbookExtensionList1.Append(workbookExtension2);

		workbook1.Append(fileVersion1);
		workbook1.Append(workbookProperties1);
		workbook1.Append(alternateContent1);
		workbook1.Append(openXmlUnknownElement1);
		workbook1.Append(bookViews1);
		workbook1.Append(sheets1);
		workbook1.Append(calculationProperties1);
		workbook1.Append(workbookExtensionList1);

		workbookPart1.Workbook = workbook1;
	}

	// Generates content of workbookStylesPart1.
	private static void GenerateWorkbookStylesPart1Content(WorkbookStylesPart workbookStylesPart1)
	{
		var stylesheet1 = new Stylesheet { MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "x14ac x16r2 xr" } };
		stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
		stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
		stylesheet1.AddNamespaceDeclaration("x16r2", "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main");
		stylesheet1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");

		var fonts1 = new Fonts
		{
			Count = (UInt32Value)1U,
			KnownFonts = true
		};

		var font1 = new Font();
		var fontSize1 = new FontSize
		{
			Val = 11D
		};
		var color1 = new Color
		{
			Theme = (UInt32Value)1U
		};
		var fontName1 = new FontName
		{
			Val = "Calibri"
		};
		var fontFamilyNumbering1 = new FontFamilyNumbering
		{
			Val = 2
		};
		var fontScheme1 = new FontScheme
		{
			Val = FontSchemeValues.Minor
		};

		font1.Append(fontSize1);
		font1.Append(color1);
		font1.Append(fontName1);
		font1.Append(fontFamilyNumbering1);
		font1.Append(fontScheme1);

		fonts1.Append(font1);

		var fills1 = new Fills
		{
			Count = (UInt32Value)2U
		};

		var fill1 = new Fill();
		var patternFill1 = new PatternFill
		{
			PatternType = PatternValues.None
		};

		fill1.Append(patternFill1);

		var fill2 = new Fill();
		var patternFill2 = new PatternFill
		{
			PatternType = PatternValues.Gray125
		};

		fill2.Append(patternFill2);

		fills1.Append(fill1);
		fills1.Append(fill2);

		var borders1 = new Borders
		{
			Count = (UInt32Value)1U
		};

		var border1 = new Border();
		var leftBorder1 = new LeftBorder();
		var rightBorder1 = new RightBorder();
		var topBorder1 = new TopBorder();
		var bottomBorder1 = new BottomBorder();
		var diagonalBorder1 = new DiagonalBorder();

		border1.Append(leftBorder1);
		border1.Append(rightBorder1);
		border1.Append(topBorder1);
		border1.Append(bottomBorder1);
		border1.Append(diagonalBorder1);

		borders1.Append(border1);

		var cellStyleFormats1 = new CellStyleFormats
		{
			Count = (UInt32Value)1U
		};
		var cellFormat1 = new CellFormat
		{
			NumberFormatId = (UInt32Value)0U,
			FontId = (UInt32Value)0U,
			FillId = (UInt32Value)0U,
			BorderId = (UInt32Value)0U
		};

		cellStyleFormats1.Append(cellFormat1);

		var cellFormats1 = new CellFormats
		{
			Count = (UInt32Value)1U
		};
		var cellFormat2 = new CellFormat
		{
			NumberFormatId = (UInt32Value)0U,
			FontId = (UInt32Value)0U,
			FillId = (UInt32Value)0U,
			BorderId = (UInt32Value)0U,
			FormatId = (UInt32Value)0U
		};

		cellFormats1.Append(cellFormat2);

		var cellStyles1 = new CellStyles { Count = (UInt32Value)1U };
		var cellStyle1 = new CellStyle { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

		cellStyles1.Append(cellStyle1);
		var differentialFormats1 = new DifferentialFormats { Count = (UInt32Value)0U };
		var tableStyles1 = new TableStyles { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

		var stylesheetExtensionList1 = new StylesheetExtensionList();

		var stylesheetExtension1 = new StylesheetExtension { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
		stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
		var slicerStyles1 = new X14.SlicerStyles { DefaultSlicerStyle = "SlicerStyleLight1" };

		stylesheetExtension1.Append(slicerStyles1);

		var stylesheetExtension2 = new StylesheetExtension { Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
		stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
		var timelineStyles1 = new X15.TimelineStyles { DefaultTimelineStyle = "TimeSlicerStyleLight1" };

		stylesheetExtension2.Append(timelineStyles1);

		stylesheetExtensionList1.Append(stylesheetExtension1);
		stylesheetExtensionList1.Append(stylesheetExtension2);

		stylesheet1.Append(fonts1);
		stylesheet1.Append(fills1);
		stylesheet1.Append(borders1);
		stylesheet1.Append(cellStyleFormats1);
		stylesheet1.Append(cellFormats1);
		stylesheet1.Append(cellStyles1);
		stylesheet1.Append(differentialFormats1);
		stylesheet1.Append(tableStyles1);
		stylesheet1.Append(stylesheetExtensionList1);

		workbookStylesPart1.Stylesheet = stylesheet1;
	}

	// Generates content of themePart1.
	private static void GenerateThemePart1Content(ThemePart themePart1)
	{
		var theme1 = new A.Theme { Name = "Office Theme" };
		theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

		var themeElements1 = new A.ThemeElements();

		var colorScheme1 = new A.ColorScheme { Name = "Office" };

		var dark1Color1 = new A.Dark1Color();
		var systemColor1 = new A.SystemColor { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

		dark1Color1.Append(systemColor1);

		var light1Color1 = new A.Light1Color();
		var systemColor2 = new A.SystemColor { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

		light1Color1.Append(systemColor2);

		var dark2Color1 = new A.Dark2Color();
		var rgbColorModelHex1 = new A.RgbColorModelHex { Val = "44546A" };

		dark2Color1.Append(rgbColorModelHex1);

		var light2Color1 = new A.Light2Color();
		var rgbColorModelHex2 = new A.RgbColorModelHex { Val = "E7E6E6" };

		light2Color1.Append(rgbColorModelHex2);

		var accent1Color1 = new A.Accent1Color();
		var rgbColorModelHex3 = new A.RgbColorModelHex { Val = "4472C4" };

		accent1Color1.Append(rgbColorModelHex3);

		var accent2Color1 = new A.Accent2Color();
		var rgbColorModelHex4 = new A.RgbColorModelHex { Val = "ED7D31" };

		accent2Color1.Append(rgbColorModelHex4);

		var accent3Color1 = new A.Accent3Color();
		var rgbColorModelHex5 = new A.RgbColorModelHex { Val = "A5A5A5" };

		accent3Color1.Append(rgbColorModelHex5);

		var accent4Color1 = new A.Accent4Color();
		var rgbColorModelHex6 = new A.RgbColorModelHex { Val = "FFC000" };

		accent4Color1.Append(rgbColorModelHex6);

		var accent5Color1 = new A.Accent5Color();
		var rgbColorModelHex7 = new A.RgbColorModelHex { Val = "5B9BD5" };

		accent5Color1.Append(rgbColorModelHex7);

		var accent6Color1 = new A.Accent6Color();
		var rgbColorModelHex8 = new A.RgbColorModelHex { Val = "70AD47" };

		accent6Color1.Append(rgbColorModelHex8);

		var hyperlink1 = new A.Hyperlink();
		var rgbColorModelHex9 = new A.RgbColorModelHex { Val = "0563C1" };

		hyperlink1.Append(rgbColorModelHex9);

		var followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
		var rgbColorModelHex10 = new A.RgbColorModelHex { Val = "954F72" };

		followedHyperlinkColor1.Append(rgbColorModelHex10);

		colorScheme1.Append(dark1Color1);
		colorScheme1.Append(light1Color1);
		colorScheme1.Append(dark2Color1);
		colorScheme1.Append(light2Color1);
		colorScheme1.Append(accent1Color1);
		colorScheme1.Append(accent2Color1);
		colorScheme1.Append(accent3Color1);
		colorScheme1.Append(accent4Color1);
		colorScheme1.Append(accent5Color1);
		colorScheme1.Append(accent6Color1);
		colorScheme1.Append(hyperlink1);
		colorScheme1.Append(followedHyperlinkColor1);

		var fontScheme2 = new A.FontScheme { Name = "Office" };

		var majorFont1 = new A.MajorFont();
		var latinFont1 = new A.LatinFont { Typeface = "Calibri Light", Panose = "020F0302020204030204" };
		var eastAsianFont1 = new A.EastAsianFont { Typeface = "" };
		var complexScriptFont1 = new A.ComplexScriptFont { Typeface = "" };
		var supplementalFont1 = new A.SupplementalFont { Script = "Jpan", Typeface = "游ゴシック Light" };
		var supplementalFont2 = new A.SupplementalFont { Script = "Hang", Typeface = "맑은 고딕" };
		var supplementalFont3 = new A.SupplementalFont { Script = "Hans", Typeface = "等线 Light" };
		var supplementalFont4 = new A.SupplementalFont { Script = "Hant", Typeface = "新細明體" };
		var supplementalFont5 = new A.SupplementalFont { Script = "Arab", Typeface = "Times New Roman" };
		var supplementalFont6 = new A.SupplementalFont { Script = "Hebr", Typeface = "Times New Roman" };
		var supplementalFont7 = new A.SupplementalFont { Script = "Thai", Typeface = "Tahoma" };
		var supplementalFont8 = new A.SupplementalFont { Script = "Ethi", Typeface = "Nyala" };
		var supplementalFont9 = new A.SupplementalFont { Script = "Beng", Typeface = "Vrinda" };
		var supplementalFont10 = new A.SupplementalFont { Script = "Gujr", Typeface = "Shruti" };
		var supplementalFont11 = new A.SupplementalFont { Script = "Khmr", Typeface = "MoolBoran" };
		var supplementalFont12 = new A.SupplementalFont { Script = "Knda", Typeface = "Tunga" };
		var supplementalFont13 = new A.SupplementalFont { Script = "Guru", Typeface = "Raavi" };
		var supplementalFont14 = new A.SupplementalFont { Script = "Cans", Typeface = "Euphemia" };
		var supplementalFont15 = new A.SupplementalFont { Script = "Cher", Typeface = "Plantagenet Cherokee" };
		var supplementalFont16 = new A.SupplementalFont { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
		var supplementalFont17 = new A.SupplementalFont { Script = "Tibt", Typeface = "Microsoft Himalaya" };
		var supplementalFont18 = new A.SupplementalFont { Script = "Thaa", Typeface = "MV Boli" };
		var supplementalFont19 = new A.SupplementalFont { Script = "Deva", Typeface = "Mangal" };
		var supplementalFont20 = new A.SupplementalFont { Script = "Telu", Typeface = "Gautami" };
		var supplementalFont21 = new A.SupplementalFont { Script = "Taml", Typeface = "Latha" };
		var supplementalFont22 = new A.SupplementalFont { Script = "Syrc", Typeface = "Estrangelo Edessa" };
		var supplementalFont23 = new A.SupplementalFont { Script = "Orya", Typeface = "Kalinga" };
		var supplementalFont24 = new A.SupplementalFont { Script = "Mlym", Typeface = "Kartika" };
		var supplementalFont25 = new A.SupplementalFont { Script = "Laoo", Typeface = "DokChampa" };
		var supplementalFont26 = new A.SupplementalFont { Script = "Sinh", Typeface = "Iskoola Pota" };
		var supplementalFont27 = new A.SupplementalFont { Script = "Mong", Typeface = "Mongolian Baiti" };
		var supplementalFont28 = new A.SupplementalFont { Script = "Viet", Typeface = "Times New Roman" };
		var supplementalFont29 = new A.SupplementalFont { Script = "Uigh", Typeface = "Microsoft Uighur" };
		var supplementalFont30 = new A.SupplementalFont { Script = "Geor", Typeface = "Sylfaen" };
		var supplementalFont31 = new A.SupplementalFont { Script = "Armn", Typeface = "Arial" };
		var supplementalFont32 = new A.SupplementalFont { Script = "Bugi", Typeface = "Leelawadee UI" };
		var supplementalFont33 = new A.SupplementalFont { Script = "Bopo", Typeface = "Microsoft JhengHei" };
		var supplementalFont34 = new A.SupplementalFont { Script = "Java", Typeface = "Javanese Text" };
		var supplementalFont35 = new A.SupplementalFont { Script = "Lisu", Typeface = "Segoe UI" };
		var supplementalFont36 = new A.SupplementalFont { Script = "Mymr", Typeface = "Myanmar Text" };
		var supplementalFont37 = new A.SupplementalFont { Script = "Nkoo", Typeface = "Ebrima" };
		var supplementalFont38 = new A.SupplementalFont { Script = "Olck", Typeface = "Nirmala UI" };
		var supplementalFont39 = new A.SupplementalFont { Script = "Osma", Typeface = "Ebrima" };
		var supplementalFont40 = new A.SupplementalFont { Script = "Phag", Typeface = "Phagspa" };
		var supplementalFont41 = new A.SupplementalFont { Script = "Syrn", Typeface = "Estrangelo Edessa" };
		var supplementalFont42 = new A.SupplementalFont { Script = "Syrj", Typeface = "Estrangelo Edessa" };
		var supplementalFont43 = new A.SupplementalFont { Script = "Syre", Typeface = "Estrangelo Edessa" };
		var supplementalFont44 = new A.SupplementalFont { Script = "Sora", Typeface = "Nirmala UI" };
		var supplementalFont45 = new A.SupplementalFont { Script = "Tale", Typeface = "Microsoft Tai Le" };
		var supplementalFont46 = new A.SupplementalFont { Script = "Talu", Typeface = "Microsoft New Tai Lue" };
		var supplementalFont47 = new A.SupplementalFont { Script = "Tfng", Typeface = "Ebrima" };

		majorFont1.Append(latinFont1);
		majorFont1.Append(eastAsianFont1);
		majorFont1.Append(complexScriptFont1);
		majorFont1.Append(supplementalFont1);
		majorFont1.Append(supplementalFont2);
		majorFont1.Append(supplementalFont3);
		majorFont1.Append(supplementalFont4);
		majorFont1.Append(supplementalFont5);
		majorFont1.Append(supplementalFont6);
		majorFont1.Append(supplementalFont7);
		majorFont1.Append(supplementalFont8);
		majorFont1.Append(supplementalFont9);
		majorFont1.Append(supplementalFont10);
		majorFont1.Append(supplementalFont11);
		majorFont1.Append(supplementalFont12);
		majorFont1.Append(supplementalFont13);
		majorFont1.Append(supplementalFont14);
		majorFont1.Append(supplementalFont15);
		majorFont1.Append(supplementalFont16);
		majorFont1.Append(supplementalFont17);
		majorFont1.Append(supplementalFont18);
		majorFont1.Append(supplementalFont19);
		majorFont1.Append(supplementalFont20);
		majorFont1.Append(supplementalFont21);
		majorFont1.Append(supplementalFont22);
		majorFont1.Append(supplementalFont23);
		majorFont1.Append(supplementalFont24);
		majorFont1.Append(supplementalFont25);
		majorFont1.Append(supplementalFont26);
		majorFont1.Append(supplementalFont27);
		majorFont1.Append(supplementalFont28);
		majorFont1.Append(supplementalFont29);
		majorFont1.Append(supplementalFont30);
		majorFont1.Append(supplementalFont31);
		majorFont1.Append(supplementalFont32);
		majorFont1.Append(supplementalFont33);
		majorFont1.Append(supplementalFont34);
		majorFont1.Append(supplementalFont35);
		majorFont1.Append(supplementalFont36);
		majorFont1.Append(supplementalFont37);
		majorFont1.Append(supplementalFont38);
		majorFont1.Append(supplementalFont39);
		majorFont1.Append(supplementalFont40);
		majorFont1.Append(supplementalFont41);
		majorFont1.Append(supplementalFont42);
		majorFont1.Append(supplementalFont43);
		majorFont1.Append(supplementalFont44);
		majorFont1.Append(supplementalFont45);
		majorFont1.Append(supplementalFont46);
		majorFont1.Append(supplementalFont47);

		var minorFont1 = new A.MinorFont();
		var latinFont2 = new A.LatinFont { Typeface = "Calibri", Panose = "020F0502020204030204" };
		var eastAsianFont2 = new A.EastAsianFont { Typeface = "" };
		var complexScriptFont2 = new A.ComplexScriptFont { Typeface = "" };
		var supplementalFont48 = new A.SupplementalFont { Script = "Jpan", Typeface = "游ゴシック" };
		var supplementalFont49 = new A.SupplementalFont { Script = "Hang", Typeface = "맑은 고딕" };
		var supplementalFont50 = new A.SupplementalFont { Script = "Hans", Typeface = "等线" };
		var supplementalFont51 = new A.SupplementalFont { Script = "Hant", Typeface = "新細明體" };
		var supplementalFont52 = new A.SupplementalFont { Script = "Arab", Typeface = "Arial" };
		var supplementalFont53 = new A.SupplementalFont { Script = "Hebr", Typeface = "Arial" };
		var supplementalFont54 = new A.SupplementalFont { Script = "Thai", Typeface = "Tahoma" };
		var supplementalFont55 = new A.SupplementalFont { Script = "Ethi", Typeface = "Nyala" };
		var supplementalFont56 = new A.SupplementalFont { Script = "Beng", Typeface = "Vrinda" };
		var supplementalFont57 = new A.SupplementalFont { Script = "Gujr", Typeface = "Shruti" };
		var supplementalFont58 = new A.SupplementalFont { Script = "Khmr", Typeface = "DaunPenh" };
		var supplementalFont59 = new A.SupplementalFont { Script = "Knda", Typeface = "Tunga" };
		var supplementalFont60 = new A.SupplementalFont { Script = "Guru", Typeface = "Raavi" };
		var supplementalFont61 = new A.SupplementalFont { Script = "Cans", Typeface = "Euphemia" };
		var supplementalFont62 = new A.SupplementalFont { Script = "Cher", Typeface = "Plantagenet Cherokee" };
		var supplementalFont63 = new A.SupplementalFont { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
		var supplementalFont64 = new A.SupplementalFont { Script = "Tibt", Typeface = "Microsoft Himalaya" };
		var supplementalFont65 = new A.SupplementalFont { Script = "Thaa", Typeface = "MV Boli" };
		var supplementalFont66 = new A.SupplementalFont { Script = "Deva", Typeface = "Mangal" };
		var supplementalFont67 = new A.SupplementalFont { Script = "Telu", Typeface = "Gautami" };
		var supplementalFont68 = new A.SupplementalFont { Script = "Taml", Typeface = "Latha" };
		var supplementalFont69 = new A.SupplementalFont { Script = "Syrc", Typeface = "Estrangelo Edessa" };
		var supplementalFont70 = new A.SupplementalFont { Script = "Orya", Typeface = "Kalinga" };
		var supplementalFont71 = new A.SupplementalFont { Script = "Mlym", Typeface = "Kartika" };
		var supplementalFont72 = new A.SupplementalFont { Script = "Laoo", Typeface = "DokChampa" };
		var supplementalFont73 = new A.SupplementalFont { Script = "Sinh", Typeface = "Iskoola Pota" };
		var supplementalFont74 = new A.SupplementalFont { Script = "Mong", Typeface = "Mongolian Baiti" };
		var supplementalFont75 = new A.SupplementalFont { Script = "Viet", Typeface = "Arial" };
		var supplementalFont76 = new A.SupplementalFont { Script = "Uigh", Typeface = "Microsoft Uighur" };
		var supplementalFont77 = new A.SupplementalFont { Script = "Geor", Typeface = "Sylfaen" };
		var supplementalFont78 = new A.SupplementalFont { Script = "Armn", Typeface = "Arial" };
		var supplementalFont79 = new A.SupplementalFont { Script = "Bugi", Typeface = "Leelawadee UI" };
		var supplementalFont80 = new A.SupplementalFont { Script = "Bopo", Typeface = "Microsoft JhengHei" };
		var supplementalFont81 = new A.SupplementalFont { Script = "Java", Typeface = "Javanese Text" };
		var supplementalFont82 = new A.SupplementalFont { Script = "Lisu", Typeface = "Segoe UI" };
		var supplementalFont83 = new A.SupplementalFont { Script = "Mymr", Typeface = "Myanmar Text" };
		var supplementalFont84 = new A.SupplementalFont { Script = "Nkoo", Typeface = "Ebrima" };
		var supplementalFont85 = new A.SupplementalFont { Script = "Olck", Typeface = "Nirmala UI" };
		var supplementalFont86 = new A.SupplementalFont { Script = "Osma", Typeface = "Ebrima" };
		var supplementalFont87 = new A.SupplementalFont { Script = "Phag", Typeface = "Phagspa" };
		var supplementalFont88 = new A.SupplementalFont { Script = "Syrn", Typeface = "Estrangelo Edessa" };
		var supplementalFont89 = new A.SupplementalFont { Script = "Syrj", Typeface = "Estrangelo Edessa" };
		var supplementalFont90 = new A.SupplementalFont { Script = "Syre", Typeface = "Estrangelo Edessa" };
		var supplementalFont91 = new A.SupplementalFont { Script = "Sora", Typeface = "Nirmala UI" };
		var supplementalFont92 = new A.SupplementalFont { Script = "Tale", Typeface = "Microsoft Tai Le" };
		var supplementalFont93 = new A.SupplementalFont { Script = "Talu", Typeface = "Microsoft New Tai Lue" };
		var supplementalFont94 = new A.SupplementalFont { Script = "Tfng", Typeface = "Ebrima" };

		minorFont1.Append(latinFont2);
		minorFont1.Append(eastAsianFont2);
		minorFont1.Append(complexScriptFont2);
		minorFont1.Append(supplementalFont48);
		minorFont1.Append(supplementalFont49);
		minorFont1.Append(supplementalFont50);
		minorFont1.Append(supplementalFont51);
		minorFont1.Append(supplementalFont52);
		minorFont1.Append(supplementalFont53);
		minorFont1.Append(supplementalFont54);
		minorFont1.Append(supplementalFont55);
		minorFont1.Append(supplementalFont56);
		minorFont1.Append(supplementalFont57);
		minorFont1.Append(supplementalFont58);
		minorFont1.Append(supplementalFont59);
		minorFont1.Append(supplementalFont60);
		minorFont1.Append(supplementalFont61);
		minorFont1.Append(supplementalFont62);
		minorFont1.Append(supplementalFont63);
		minorFont1.Append(supplementalFont64);
		minorFont1.Append(supplementalFont65);
		minorFont1.Append(supplementalFont66);
		minorFont1.Append(supplementalFont67);
		minorFont1.Append(supplementalFont68);
		minorFont1.Append(supplementalFont69);
		minorFont1.Append(supplementalFont70);
		minorFont1.Append(supplementalFont71);
		minorFont1.Append(supplementalFont72);
		minorFont1.Append(supplementalFont73);
		minorFont1.Append(supplementalFont74);
		minorFont1.Append(supplementalFont75);
		minorFont1.Append(supplementalFont76);
		minorFont1.Append(supplementalFont77);
		minorFont1.Append(supplementalFont78);
		minorFont1.Append(supplementalFont79);
		minorFont1.Append(supplementalFont80);
		minorFont1.Append(supplementalFont81);
		minorFont1.Append(supplementalFont82);
		minorFont1.Append(supplementalFont83);
		minorFont1.Append(supplementalFont84);
		minorFont1.Append(supplementalFont85);
		minorFont1.Append(supplementalFont86);
		minorFont1.Append(supplementalFont87);
		minorFont1.Append(supplementalFont88);
		minorFont1.Append(supplementalFont89);
		minorFont1.Append(supplementalFont90);
		minorFont1.Append(supplementalFont91);
		minorFont1.Append(supplementalFont92);
		minorFont1.Append(supplementalFont93);
		minorFont1.Append(supplementalFont94);

		fontScheme2.Append(majorFont1);
		fontScheme2.Append(minorFont1);

		var formatScheme1 = new A.FormatScheme { Name = "Office" };

		var fillStyleList1 = new A.FillStyleList();

		var solidFill1 = new A.SolidFill();
		var schemeColor1 = new A.SchemeColor { Val = A.SchemeColorValues.PhColor };

		solidFill1.Append(schemeColor1);

		var gradientFill1 = new A.GradientFill { RotateWithShape = true };

		var gradientStopList1 = new A.GradientStopList();

		var gradientStop1 = new A.GradientStop { Position = 0 };

		var schemeColor2 = new A.SchemeColor { Val = A.SchemeColorValues.PhColor };
		var luminanceModulation1 = new A.LuminanceModulation { Val = 110000 };
		var saturationModulation1 = new A.SaturationModulation { Val = 105000 };
		var tint1 = new A.Tint { Val = 67000 };

		schemeColor2.Append(luminanceModulation1);
		schemeColor2.Append(saturationModulation1);
		schemeColor2.Append(tint1);

		gradientStop1.Append(schemeColor2);

		var gradientStop2 = new A.GradientStop { Position = 50000 };

		var schemeColor3 = new A.SchemeColor { Val = A.SchemeColorValues.PhColor };
		var luminanceModulation2 = new A.LuminanceModulation { Val = 105000 };
		var saturationModulation2 = new A.SaturationModulation { Val = 103000 };
		var tint2 = new A.Tint { Val = 73000 };

		schemeColor3.Append(luminanceModulation2);
		schemeColor3.Append(saturationModulation2);
		schemeColor3.Append(tint2);

		gradientStop2.Append(schemeColor3);

		var gradientStop3 = new A.GradientStop { Position = 100000 };

		var schemeColor4 = new A.SchemeColor { Val = A.SchemeColorValues.PhColor };
		var luminanceModulation3 = new A.LuminanceModulation { Val = 105000 };
		var saturationModulation3 = new A.SaturationModulation { Val = 109000 };
		var tint3 = new A.Tint { Val = 81000 };

		schemeColor4.Append(luminanceModulation3);
		schemeColor4.Append(saturationModulation3);
		schemeColor4.Append(tint3);

		gradientStop3.Append(schemeColor4);

		gradientStopList1.Append(gradientStop1);
		gradientStopList1.Append(gradientStop2);
		gradientStopList1.Append(gradientStop3);
		var linearGradientFill1 = new A.LinearGradientFill { Angle = 5400000, Scaled = false };

		gradientFill1.Append(gradientStopList1);
		gradientFill1.Append(linearGradientFill1);

		var gradientFill2 = new A.GradientFill { RotateWithShape = true };

		var gradientStopList2 = new A.GradientStopList();

		var gradientStop4 = new A.GradientStop { Position = 0 };

		var schemeColor5 = new A.SchemeColor { Val = A.SchemeColorValues.PhColor };
		var saturationModulation4 = new A.SaturationModulation { Val = 103000 };
		var luminanceModulation4 = new A.LuminanceModulation { Val = 102000 };
		var tint4 = new A.Tint { Val = 94000 };

		schemeColor5.Append(saturationModulation4);
		schemeColor5.Append(luminanceModulation4);
		schemeColor5.Append(tint4);

		gradientStop4.Append(schemeColor5);

		var gradientStop5 = new A.GradientStop { Position = 50000 };

		var schemeColor6 = new A.SchemeColor { Val = A.SchemeColorValues.PhColor };
		var saturationModulation5 = new A.SaturationModulation { Val = 110000 };
		var luminanceModulation5 = new A.LuminanceModulation { Val = 100000 };
		var shade1 = new A.Shade { Val = 100000 };

		schemeColor6.Append(saturationModulation5);
		schemeColor6.Append(luminanceModulation5);
		schemeColor6.Append(shade1);

		gradientStop5.Append(schemeColor6);

		var gradientStop6 = new A.GradientStop { Position = 100000 };

		var schemeColor7 = new A.SchemeColor { Val = A.SchemeColorValues.PhColor };
		var luminanceModulation6 = new A.LuminanceModulation { Val = 99000 };
		var saturationModulation6 = new A.SaturationModulation { Val = 120000 };
		var shade2 = new A.Shade { Val = 78000 };

		schemeColor7.Append(luminanceModulation6);
		schemeColor7.Append(saturationModulation6);
		schemeColor7.Append(shade2);

		gradientStop6.Append(schemeColor7);

		gradientStopList2.Append(gradientStop4);
		gradientStopList2.Append(gradientStop5);
		gradientStopList2.Append(gradientStop6);
		var linearGradientFill2 = new A.LinearGradientFill { Angle = 5400000, Scaled = false };

		gradientFill2.Append(gradientStopList2);
		gradientFill2.Append(linearGradientFill2);

		fillStyleList1.Append(solidFill1);
		fillStyleList1.Append(gradientFill1);
		fillStyleList1.Append(gradientFill2);

		var lineStyleList1 = new A.LineStyleList();

		var outline1 = new A.Outline { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

		var solidFill2 = new A.SolidFill();
		var schemeColor8 = new A.SchemeColor { Val = A.SchemeColorValues.PhColor };

		solidFill2.Append(schemeColor8);
		var presetDash1 = new A.PresetDash { Val = A.PresetLineDashValues.Solid };
		var miter1 = new A.Miter { Limit = 800000 };

		outline1.Append(solidFill2);
		outline1.Append(presetDash1);
		outline1.Append(miter1);

		var outline2 = new A.Outline { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

		var solidFill3 = new A.SolidFill();
		var schemeColor9 = new A.SchemeColor { Val = A.SchemeColorValues.PhColor };

		solidFill3.Append(schemeColor9);
		var presetDash2 = new A.PresetDash { Val = A.PresetLineDashValues.Solid };
		var miter2 = new A.Miter { Limit = 800000 };

		outline2.Append(solidFill3);
		outline2.Append(presetDash2);
		outline2.Append(miter2);

		var outline3 = new A.Outline { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

		var solidFill4 = new A.SolidFill();
		var schemeColor10 = new A.SchemeColor { Val = A.SchemeColorValues.PhColor };

		solidFill4.Append(schemeColor10);
		var presetDash3 = new A.PresetDash { Val = A.PresetLineDashValues.Solid };
		var miter3 = new A.Miter { Limit = 800000 };

		outline3.Append(solidFill4);
		outline3.Append(presetDash3);
		outline3.Append(miter3);

		lineStyleList1.Append(outline1);
		lineStyleList1.Append(outline2);
		lineStyleList1.Append(outline3);

		var effectStyleList1 = new A.EffectStyleList();

		var effectStyle1 = new A.EffectStyle();
		var effectList1 = new A.EffectList();

		effectStyle1.Append(effectList1);

		var effectStyle2 = new A.EffectStyle();
		var effectList2 = new A.EffectList();

		effectStyle2.Append(effectList2);

		var effectStyle3 = new A.EffectStyle();

		var effectList3 = new A.EffectList();

		var outerShadow1 = new A.OuterShadow { BlurRadius = 57150L, Distance = 19050L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

		var rgbColorModelHex11 = new A.RgbColorModelHex { Val = "000000" };
		var alpha1 = new A.Alpha { Val = 63000 };

		rgbColorModelHex11.Append(alpha1);

		outerShadow1.Append(rgbColorModelHex11);

		effectList3.Append(outerShadow1);

		effectStyle3.Append(effectList3);

		effectStyleList1.Append(effectStyle1);
		effectStyleList1.Append(effectStyle2);
		effectStyleList1.Append(effectStyle3);

		var backgroundFillStyleList1 = new A.BackgroundFillStyleList();

		var solidFill5 = new A.SolidFill();
		var schemeColor11 = new A.SchemeColor { Val = A.SchemeColorValues.PhColor };

		solidFill5.Append(schemeColor11);

		var solidFill6 = new A.SolidFill();

		var schemeColor12 = new A.SchemeColor { Val = A.SchemeColorValues.PhColor };
		var tint5 = new A.Tint { Val = 95000 };
		var saturationModulation7 = new A.SaturationModulation { Val = 170000 };

		schemeColor12.Append(tint5);
		schemeColor12.Append(saturationModulation7);

		solidFill6.Append(schemeColor12);

		var gradientFill3 = new A.GradientFill { RotateWithShape = true };

		var gradientStopList3 = new A.GradientStopList();

		var gradientStop7 = new A.GradientStop { Position = 0 };

		var schemeColor13 = new A.SchemeColor { Val = A.SchemeColorValues.PhColor };
		var tint6 = new A.Tint { Val = 93000 };
		var saturationModulation8 = new A.SaturationModulation { Val = 150000 };
		var shade3 = new A.Shade { Val = 98000 };
		var luminanceModulation7 = new A.LuminanceModulation { Val = 102000 };

		schemeColor13.Append(tint6);
		schemeColor13.Append(saturationModulation8);
		schemeColor13.Append(shade3);
		schemeColor13.Append(luminanceModulation7);

		gradientStop7.Append(schemeColor13);

		var gradientStop8 = new A.GradientStop { Position = 50000 };

		var schemeColor14 = new A.SchemeColor { Val = A.SchemeColorValues.PhColor };
		var tint7 = new A.Tint { Val = 98000 };
		var saturationModulation9 = new A.SaturationModulation { Val = 130000 };
		var shade4 = new A.Shade { Val = 90000 };
		var luminanceModulation8 = new A.LuminanceModulation { Val = 103000 };

		schemeColor14.Append(tint7);
		schemeColor14.Append(saturationModulation9);
		schemeColor14.Append(shade4);
		schemeColor14.Append(luminanceModulation8);

		gradientStop8.Append(schemeColor14);

		var gradientStop9 = new A.GradientStop { Position = 100000 };

		var schemeColor15 = new A.SchemeColor { Val = A.SchemeColorValues.PhColor };
		var shade5 = new A.Shade { Val = 63000 };
		var saturationModulation10 = new A.SaturationModulation { Val = 120000 };

		schemeColor15.Append(shade5);
		schemeColor15.Append(saturationModulation10);

		gradientStop9.Append(schemeColor15);

		gradientStopList3.Append(gradientStop7);
		gradientStopList3.Append(gradientStop8);
		gradientStopList3.Append(gradientStop9);
		var linearGradientFill3 = new A.LinearGradientFill { Angle = 5400000, Scaled = false };

		gradientFill3.Append(gradientStopList3);
		gradientFill3.Append(linearGradientFill3);

		backgroundFillStyleList1.Append(solidFill5);
		backgroundFillStyleList1.Append(solidFill6);
		backgroundFillStyleList1.Append(gradientFill3);

		formatScheme1.Append(fillStyleList1);
		formatScheme1.Append(lineStyleList1);
		formatScheme1.Append(effectStyleList1);
		formatScheme1.Append(backgroundFillStyleList1);

		themeElements1.Append(colorScheme1);
		themeElements1.Append(fontScheme2);
		themeElements1.Append(formatScheme1);
		var objectDefaults1 = new A.ObjectDefaults();
		var extraColorSchemeList1 = new A.ExtraColorSchemeList();

		var officeStyleSheetExtensionList1 = new A.OfficeStyleSheetExtensionList();

		var officeStyleSheetExtension1 = new A.OfficeStyleSheetExtension { Uri = "{05A4C25C-085E-4340-85A3-A5531E510DB2}" };

		var themeFamily1 = new Thm15.ThemeFamily { Name = "Office Theme", Id = "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}", Vid = "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}" };
		themeFamily1.AddNamespaceDeclaration("thm15", "http://schemas.microsoft.com/office/thememl/2012/main");

		officeStyleSheetExtension1.Append(themeFamily1);

		officeStyleSheetExtensionList1.Append(officeStyleSheetExtension1);

		theme1.Append(themeElements1);
		theme1.Append(objectDefaults1);
		theme1.Append(extraColorSchemeList1);
		theme1.Append(officeStyleSheetExtensionList1);

		themePart1.Theme = theme1;
	}

	// Generates content of worksheetPart1.
	private static void GenerateWorksheetPart1Content(WorksheetPart worksheetPart1)
	{
		var worksheet1 = new Worksheet { MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "x14ac xr xr2 xr3" } };
		worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
		worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
		worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
		worksheet1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
		worksheet1.AddNamespaceDeclaration("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");
		worksheet1.AddNamespaceDeclaration("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3");
		worksheet1.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{C2C8BF20-22AB-430D-82EB-7D63C9600358}"));
		var sheetDimension1 = new SheetDimension { Reference = "B2:C3" };

		var sheetViews1 = new SheetViews();

		var sheetView1 = new SheetView { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
		var selection1 = new Selection { ActiveCell = "B2", SequenceOfReferences = new ListValue<StringValue> { InnerText = "B2:C3" } };

		sheetView1.Append(selection1);

		sheetViews1.Append(sheetView1);
		var sheetFormatProperties1 = new SheetFormatProperties { DefaultRowHeight = 15D, DyDescent = 0.25D };

		var columns1 = new Columns();
		var column1 = new Column { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 9.5703125D, CustomWidth = true };

		columns1.Append(column1);

		var sheetData1 = new SheetData();

		var row1 = new Row { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue> { InnerText = "2:3" }, DyDescent = 0.25D };

		var cell1 = new Cell { CellReference = "B2", DataType = CellValues.SharedString };
		var cellValue1 = new CellValue
		{
			Text = "0"
		};

		cell1.Append(cellValue1);

		var cell2 = new Cell { CellReference = "C2", DataType = CellValues.SharedString };
		var cellValue2 = new CellValue
		{
			Text = "1"
		};

		cell2.Append(cellValue2);

		row1.Append(cell1);
		row1.Append(cell2);

		var row2 = new Row { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue> { InnerText = "2:3" }, DyDescent = 0.25D };

		var cell3 = new Cell { CellReference = "B3", DataType = CellValues.SharedString };
		var cellValue3 = new CellValue
		{
			Text = "2"
		};

		cell3.Append(cellValue3);

		var cell4 = new Cell { CellReference = "C3", DataType = CellValues.SharedString };
		var cellValue4 = new CellValue
		{
			Text = "3"
		};

		cell4.Append(cellValue4);

		row2.Append(cell3);
		row2.Append(cell4);

		sheetData1.Append(row1);
		sheetData1.Append(row2);
		var pageMargins1 = new PageMargins { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };

		var tableParts1 = new TableParts { Count = (UInt32Value)1U };
		var tablePart1 = new TablePart { Id = "rId1" };

		tableParts1.Append(tablePart1);

		worksheet1.Append(sheetDimension1);
		worksheet1.Append(sheetViews1);
		worksheet1.Append(sheetFormatProperties1);
		worksheet1.Append(columns1);
		worksheet1.Append(sheetData1);
		worksheet1.Append(pageMargins1);
		worksheet1.Append(tableParts1);

		worksheetPart1.Worksheet = worksheet1;
	}

	// Generates content of tableDefinitionPart1.
	private static void GenerateTableDefinitionPart1Content(TableDefinitionPart tableDefinitionPart1)
	{
		var table1 = new Table { Id = (UInt32Value)1U, Name = "Table1", DisplayName = "Table1", Reference = "B2:C3", TotalsRowShown = false, MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "xr xr3" } };
		table1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
		table1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
		table1.AddNamespaceDeclaration("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3");
		table1.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{F5046044-E24F-4E50-B026-CF2785460B75}"));

		var autoFilter1 = new AutoFilter { Reference = "B2:C3" };
		autoFilter1.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{668AE273-0804-43E9-A918-462EC2855C7E}"));

		var tableColumns1 = new TableColumns { Count = (UInt32Value)2U };

		var tableColumn1 = new TableColumn { Id = (UInt32Value)1U, Name = "AAAAA" };
		tableColumn1.SetAttribute(new OpenXmlAttribute("xr3", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3", "{AD06AF15-8079-4AA5-9138-801B1A00364F}"));

		var tableColumn2 = new TableColumn { Id = (UInt32Value)2U, Name = "BBBBB" };
		tableColumn2.SetAttribute(new OpenXmlAttribute("xr3", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3", "{8615FFE0-D5EC-4846-AD63-7776E89979EA}"));

		tableColumns1.Append(tableColumn1);
		tableColumns1.Append(tableColumn2);
		var tableStyleInfo1 = new TableStyleInfo { Name = "TableStyleMedium14", ShowFirstColumn = false, ShowLastColumn = false, ShowRowStripes = true, ShowColumnStripes = false };

		table1.Append(autoFilter1);
		table1.Append(tableColumns1);
		table1.Append(tableStyleInfo1);

		tableDefinitionPart1.Table = table1;
	}

	// Generates content of sharedStringTablePart1.
	private static void GenerateSharedStringTablePart1Content(SharedStringTablePart sharedStringTablePart1)
	{
		var sharedStringTable1 = new SharedStringTable { Count = (UInt32Value)4U, UniqueCount = (UInt32Value)4U };

		var sharedStringItem1 = new SharedStringItem();
		var text1 = new Text
		{
			Text = "AAAAA"
		};

		sharedStringItem1.Append(text1);

		var sharedStringItem2 = new SharedStringItem();
		var text2 = new Text
		{
			Text = "BBBBB"
		};

		sharedStringItem2.Append(text2);

		var sharedStringItem3 = new SharedStringItem();
		var text3 = new Text
		{
			Text = "CCCCC"
		};

		sharedStringItem3.Append(text3);

		var sharedStringItem4 = new SharedStringItem();
		var text4 = new Text
		{
			Text = "DDDDD"
		};

		sharedStringItem4.Append(text4);

		sharedStringTable1.Append(sharedStringItem1);
		sharedStringTable1.Append(sharedStringItem2);
		sharedStringTable1.Append(sharedStringItem3);
		sharedStringTable1.Append(sharedStringItem4);

		sharedStringTablePart1.SharedStringTable = sharedStringTable1;
	}

	private static void SetPackageProperties(OpenXmlPackage document)
	{
		document.PackageProperties.Creator = "david.bond";
		document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2020-05-12T12:07:26Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
		document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2020-05-12T12:08:11Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
		document.PackageProperties.LastModifiedBy = "david.bond";
	}
}
