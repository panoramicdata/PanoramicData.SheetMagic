namespace PanoramicData.SheetMagic.Test;

public class HeaderTests : Test
{
	[Fact]
	public void HeaderAsNumbers_Succeeds()
	{
		List<Extended<object>> items;
		using (var magicSpreadsheet = new MagicSpreadsheet(GetSheetFileInfo("HeaderTest")))
		{
			magicSpreadsheet.Load();
			items = magicSpreadsheet.GetExtendedList<object>(magicSpreadsheet.SheetNames.FirstOrDefault() ?? string.Empty);
		}

		_ = items.Should().NotBeNull();
	}
}