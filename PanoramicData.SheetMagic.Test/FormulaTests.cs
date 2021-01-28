using Xunit;

namespace PanoramicData.SheetMagic.Test
{
	public class FormulaTests : Test
	{
		[Fact]
		public void GetValueFromFormula_Succeeds()
		{
			using var magicSpreadsheet = new MagicSpreadsheet(GetSheetFileInfo("FormulaTest"), new Options { StopProcessingOnFirstEmptyRow = true });
			magicSpreadsheet.Load();
			var items = magicSpreadsheet.GetExtendedList<object>("Sheet2");

			// Check the items
			Assert.NotEmpty(items);
			Assert.True(items.Count > 0);
			Assert.Equal("6", items[0].Properties["Total"]);
		}
	}
}
