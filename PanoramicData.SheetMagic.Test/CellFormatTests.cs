using AwesomeAssertions;
using Xunit;

namespace PanoramicData.SheetMagic.Test;

public class CellFormatTests : Test
{
	[Fact]
	public void CheckFormats_Succeeds()
	{
		using var magicSpreadsheet = new MagicSpreadsheet(GetSheetFileInfo("Cell Formats"), new Options { StopProcessingOnFirstEmptyRow = true });
		magicSpreadsheet.Load();
		var items = magicSpreadsheet.GetExtendedList<object>("Sheet1");

		// Check the items
		Assert.NotEmpty(items);
		(items.Count > 0).Should().BeTrue();
		items[0].Properties["General"].Should().Be("Happy Christmas!");
		Assert.Equal("99.00", items[0].Properties["Number (N2)"]);
		Assert.Equal("99.0", items[0].Properties["Number (N1)"]);
		Assert.Equal("99", items[0].Properties["Number (N0)"]);
		Assert.Equal("25/05/1975", items[0].Properties["Date"]);
		Assert.Equal("50.0%", items[0].Properties["Percentage N1"]);
		Assert.Equal("50.00%", items[0].Properties["Percentage N2"]);
		Assert.Equal("Here is some text", items[0].Properties["Text"]);
	}

	[Fact]
	public void ExtendedListVariousOmittedEmptyCells_Succeeds()
	{
		// The input file has 8 cell XML elements in some rows, and 9 in others.
		// Excel omits cells e.g. if no formatting, or empty, and various combinations.
		// We should expect the same number of items regardless
		// See MS-21227

		using var magicSpreadsheet =
			new MagicSpreadsheet(GetSheetFileInfo("XML Missing Some Empty String Cells"),
			new Options
			{
				LoadNullExtendedProperties = true,
				StopProcessingOnFirstEmptyRow = true
			});
		magicSpreadsheet.Load();

		var items = magicSpreadsheet.GetExtendedList<object>("Sheet1");

		// Check the items
		Assert.NotEmpty(items);
		Assert.True(items.Count == 5);
		Assert.Equal(items[0].Properties.Count, 9);
		Assert.Equal(items[1].Properties.Count, 9);
		Assert.Equal(items[2].Properties.Count, 9);
		Assert.Equal(items[3].Properties.Count, 9);
		Assert.Equal(items[4].Properties.Count, 9);
	}
}