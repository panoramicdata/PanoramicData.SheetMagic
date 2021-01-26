using System;
using System.IO;
using System.Reflection;

namespace PanoramicData.SheetMagic.Test
{
	public abstract class Test
	{
		protected static FileInfo GetXlsxTempFileInfo()
	 		=> new FileInfo(Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx"));

		protected static FileInfo GetSheetFileInfo(string worksheetName)
		{
			var location = typeof(LoadSheetTests).GetTypeInfo().Assembly.Location;
			var dirPath = Path.Combine(Path.GetDirectoryName(location)!, "../../../Sheets");
			return new FileInfo(Path.Combine(dirPath, $"{worksheetName}.xlsx"));
		}
	}
}
