using AwesomeAssertions;
using System.Collections.Generic;
using Xunit;

namespace PanoramicData.SheetMagic.Test;

public class SheetNameTests : Test
{
	[Fact]
	public void SheetNames_Succeeds()
	{
		// Load the parent/child relationships
		List<string> sheetNames;
		using (var magicSpreadsheet = new MagicSpreadsheet(GetSheetFileInfo("SitesAndDevices")))
		{
			magicSpreadsheet.Load();
			sheetNames = magicSpreadsheet.SheetNames;
		}

		// Loaded
		_ = sheetNames.Should().HaveCount(2);
		_ = sheetNames[0].Should().NotBeNull();
		_ = sheetNames[0].Should().Be("Sites");
		_ = sheetNames[1].Should().NotBeNull();
		_ = sheetNames[1].Should().Be("Devices");
	}
}