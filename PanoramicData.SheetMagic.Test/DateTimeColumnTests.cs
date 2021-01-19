using System.IO;
using System.Reflection;
using Xunit;

namespace PanoramicData.SheetMagic.Test
{
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
		}

		private static FileInfo GetSheetFileInfo(string worksheetName)
		{
			var location = typeof(LoadSheetTests).GetTypeInfo().Assembly.Location;
			var dirPath = Path.Combine(Path.GetDirectoryName(location)!, "../../../Sheets");
			return new FileInfo(Path.Combine(dirPath, $"{worksheetName}.xlsx"));
		}
	}
}