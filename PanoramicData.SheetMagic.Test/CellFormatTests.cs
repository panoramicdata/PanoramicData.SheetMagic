using Xunit;

namespace PanoramicData.SheetMagic.Test
{
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
			Assert.True(items.Count > 0);
			Assert.Equal("Happy Christmas!", items[0].Properties["General"]);
			Assert.Equal("99.00", items[0].Properties["Number (N2)"]);
			Assert.Equal("99.0", items[0].Properties["Number (N1)"]);
			Assert.Equal("99", items[0].Properties["Number (N0)"]);
			Assert.Equal("25/05/1975", items[0].Properties["Date"]);
			Assert.Equal("50.0%", items[0].Properties["Percentage N1"]);
			Assert.Equal("50.00%", items[0].Properties["Percentage N2"]);
			Assert.Equal("Here is some text", items[0].Properties["Text"]);
		}
	}
}