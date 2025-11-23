using AwesomeAssertions;
using Xunit;

namespace PanoramicData.SheetMagic.Test;

public class CellFormatTests : Test
{
	private const int ExpectedPropertyCount = 9;
	private const int ExpectedItemCount = 5;

	[Fact]
	public void CheckFormats_Succeeds()
	{
		using var magicSpreadsheet = new MagicSpreadsheet(GetSheetFileInfo("Cell Formats"), new Options { StopProcessingOnFirstEmptyRow = true });
		magicSpreadsheet.Load();
		var items = magicSpreadsheet.GetExtendedList<object>("Sheet1");

		// Check the items
		items.Should().NotBeNullOrEmpty();
		(items.Count > 0).Should().BeTrue();
		items[0].Properties["General"].Should().Be("Happy Christmas!");
		items[0].Properties["Number (N2)"].Should().Be("99.00");
		items[0].Properties["Number (N1)"].Should().Be("99.0");
		items[0].Properties["Number (N0)"].Should().Be("99");
		items[0].Properties["Date"].Should().Be("25/05/1975");
		items[0].Properties["Percentage N1"].Should().Be("50.0%");
		items[0].Properties["Percentage N2"].Should().Be("50.00%");
		items[0].Properties["Text"].Should().Be("Here is some text");
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
		items.Should().NotBeNullOrEmpty();

		items.Should().HaveCount(ExpectedItemCount);

		items[0].Properties.Should().HaveCount(ExpectedPropertyCount);
		items[1].Properties.Should().HaveCount(ExpectedPropertyCount);
		items[2].Properties.Should().HaveCount(ExpectedPropertyCount);
		items[3].Properties.Should().HaveCount(ExpectedPropertyCount);
		items[4].Properties.Should().HaveCount(ExpectedPropertyCount);
	}
}