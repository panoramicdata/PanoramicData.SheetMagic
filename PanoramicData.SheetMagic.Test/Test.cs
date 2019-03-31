using System;
using System.IO;

namespace PanoramicData.SheetMagic.Test
{
	public abstract class Test
	{
		protected static FileInfo GetXlsxTempFileInfo()
			=> new FileInfo(Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx"));
	}
}
