﻿using PanoramicData.SheetMagic.Test.Models;
using System;
using System.Collections.Generic;
using Xunit;

namespace PanoramicData.SheetMagic.Test
{
	public class AddSheet : Test
	{
		[Theory]
		[InlineData("12345678901234567890123456789012")]
		[InlineData("abcdefghijklmnopqrstuvwxyz123456")]
		public void AddSheet_SheetNameTooLong_Fails(string badSheetName)
		{
			var fileInfo = GetXlsxTempFileInfo();

			try
			{
				using var s = new MagicSpreadsheet(fileInfo);
				Assert.Throws<ArgumentException>(() => s.AddSheet(new List<SimpleAnimal>(), badSheetName));
			}
			finally
			{
				fileInfo.Delete();
			}
		}

		[Fact]
		public void AddSheet_SheetNameAlreadyExists_Fails()
		{
			var fileInfo = GetXlsxTempFileInfo();

			try
			{
				using var s = new MagicSpreadsheet(fileInfo);
				s.AddSheet(new List<SimpleAnimal>(), "Sheet1");
				Assert.Throws<ArgumentException>(() => s.AddSheet(new List<SimpleAnimal>(), "Sheet1"));
			}
			finally
			{
				fileInfo.Delete();
			}
		}

		[Fact]
		public void AddSheet_SheetWithExtraExtendedProperties_Succeeds()
		{
			var fileInfo = GetXlsxTempFileInfo();

			try
			{
				using var s = new MagicSpreadsheet(fileInfo);

				var sheetOptions = new AddSheetOptions
				{
					TableOptions = new TableOptions { XlsxTableStyle = XlsxTableStyle.TableStyleDark1 }
				};

				s.AddSheet(new List<Extended<object>>
					 {
						  new Extended<object>(new object(), new Dictionary<string, object?> { { "Id", 10 }, { "My Name", "Ryan" } })
					 }, "Subscriptions", sheetOptions);
				s.Save();
			}
			finally
			{
				fileInfo.Delete();
			}
		}
	}
}
