using PanoramicData.SheetMagic.Test.Models;
using System;
using System.Collections.Generic;
using System.IO;
using Xunit;

namespace PanoramicData.SheetMagic.Test
{
	public class AddSheet
	{
		[Theory]
		[InlineData("12345678901234567890123456789012")]
		[InlineData("abcdefghijklmnopqrstuvwxyz123456")]
		public void AddSheet_SheetNameTooLong_Fails(string badSheetName)
		{
			var fileInfo = new FileInfo(Path.GetTempFileName());

			try
			{
				using (var s = new MagicSpreadsheet(fileInfo))
				{
					Assert.Throws<ArgumentException>(() => s.AddSheet(new List<SimpleAnimal>(), badSheetName));
				}
			}
			finally
			{
				fileInfo.Delete();
			}
		}

		[Fact]
		public void AddSheet_SheetNameAlreadyExists_Fails()
		{
			var fileInfo = new FileInfo(Path.GetTempFileName());

			try
			{
				using (var s = new MagicSpreadsheet(fileInfo))
				{
					s.AddSheet(new List<SimpleAnimal>(), "Sheet1");
					Assert.Throws<ArgumentException>(() => s.AddSheet(new List<SimpleAnimal>(), "Sheet1"));
				}
			}
			finally
			{
				fileInfo.Delete();
			}
		}
	}
}
