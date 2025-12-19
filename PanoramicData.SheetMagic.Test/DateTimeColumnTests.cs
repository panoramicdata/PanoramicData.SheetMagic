namespace PanoramicData.SheetMagic.Test;

public class DateTimeColumnTests : Test
{
	[Fact]
	public void GetValuesAsDateTime_Succeeds()
	{
		// DateTimeTest and DateTimeTest2 have a couple of date formats in (and number formats, etc)
		using var magicSpreadsheet = new MagicSpreadsheet(GetSheetFileInfo("DateTimeTest2"), new Options { StopProcessingOnFirstEmptyRow = true });
		magicSpreadsheet.Load();
		var items = magicSpreadsheet.GetExtendedList<object>("Sheet1");

		// Loaded
		_ = items.Should().NotBeNull();
		_ = items.Should().HaveCountGreaterThan(0);
	}
}