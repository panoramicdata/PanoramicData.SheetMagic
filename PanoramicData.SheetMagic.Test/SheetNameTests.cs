using FluentAssertions;
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
		sheetNames.Count.Should().Be(2);
		sheetNames[0].Should().NotBeNull();
		sheetNames[0].Should().Be("Sites");
		sheetNames[1].Should().NotBeNull();
		sheetNames[1].Should().Be("Devices");
	}
}