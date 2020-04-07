using FluentAssertions;
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
			var a = new Extended<object>(
				new object(),
				new Dictionary<string, object?>()
				{
					{ "a", "b" }
				}
			);
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
				b.Should().NotBeNullOrEmpty();
				var firstItem = b[0];
				firstItem.Properties.Keys.Should().Contain("a");
				firstItem.Properties["a"].Should().Be("b");
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
			var car = new Extended<Car>(new Car
			{
				Id = 1,
				Name = "Yumyum",
				WheelCount = 4,
				WeightKg = carWeightKg
			},
				new Dictionary<string, object?>
				{
					{ customPropertyName, customPropertyValue }
				});
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
				cars.Should().NotBeNullOrEmpty();
				var firstCar = cars[0];
				firstCar.Item.Should().NotBeNull();
				carWeightKg.Should().Be(firstCar.Item.WeightKg);
				firstCar.Properties.Keys.Should().Contain(customPropertyName);
				firstCar.Properties[customPropertyName].Should().Be(customPropertyValue);
			}
			finally
			{
				fileInfo.Delete();
			}
		}
	}
}