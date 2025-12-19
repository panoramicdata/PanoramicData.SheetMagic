namespace PanoramicData.SheetMagic.Test;

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
		(items.Count > 0).Should().BeTrue();
		items[0].Properties["Total"].Should().Be("6");
	}
}
