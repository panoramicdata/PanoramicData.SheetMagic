using PanoramicData.SheetMagic.Test.Models;
using System.Collections.Generic;
using Xunit;

namespace PanoramicData.SheetMagic.Test
{
	public class SaveSheet : Test
	{
		[Fact]
		public void SaveSheet_WithNoData_Succeeds()
		{
			var fileInfo = GetXlsxTempFileInfo();

			try
			{
				using var s = new MagicSpreadsheet(fileInfo);
				s.Save();
			}
			finally
			{
				fileInfo.Delete();
			}
		}

		[Fact]
		public void SavingWithExtendedObject_Succeeds()
		{
			var a = new Extended<object>
			{
				Properties = new Dictionary<string, object>()
				{
					{ "a", "b" }
				}
			};
			var fileInfo = GetXlsxTempFileInfo();

			try
			{
				// Save
				using (var s1 = new MagicSpreadsheet(fileInfo))
				{
					s1.AddSheet(new List<Extended<object>> { a });
					s1.Save();
				}

				using var s2 = new MagicSpreadsheet(fileInfo);
				s2.Load();
				var b = s2.GetExtendedList<object>();
				Assert.NotNull(b);
				Assert.NotEmpty(b);
				var firstItem = b[0];
				Assert.True(firstItem.Properties.ContainsKey("a"));
				Assert.Equal("b", firstItem.Properties["a"]);
			}
			finally
			{
				fileInfo.Delete();
			}
		}

		[Fact]
		public void SavingWithExtendedModel_Succeeds()
		{
			const int carWeightKg = 2200;
			const string customPropertyName = "CustomPropertyName";
			const string customPropertyValue = "CustomPropertyValue";
			var car = new Extended<Car>
			{
				Item = new Car
				{
					Id = 1,
					Name = "Yumyum",
					WheelCount = 4,
					WeightKg = carWeightKg
				},
				Properties = new Dictionary<string, object>()
				{
					{ customPropertyName, customPropertyValue }
				}
			};
			var fileInfo = GetXlsxTempFileInfo();

			try
			{
				// Save
				using (var s1 = new MagicSpreadsheet(fileInfo))
				{
					s1.AddSheet(new List<Extended<Car>> { car });
					s1.Save();
				}

				using var s2 = new MagicSpreadsheet(fileInfo);
				s2.Load();
				var cars = s2.GetExtendedList<Car>();
				Assert.NotNull(cars);
				Assert.NotEmpty(cars);
				var firstCar = cars[0];
				Assert.NotNull(firstCar.Item);
				Assert.Equal(carWeightKg, firstCar.Item.WeightKg);
				Assert.True(firstCar.Properties.ContainsKey(customPropertyName));
				Assert.Equal(customPropertyValue, firstCar.Properties[customPropertyName]);
			}
			finally
			{
				fileInfo.Delete();
			}
		}
	}
}