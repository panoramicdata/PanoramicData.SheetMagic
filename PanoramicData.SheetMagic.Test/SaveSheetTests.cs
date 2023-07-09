using FluentAssertions;
using PanoramicData.SheetMagic.Test.Models;
using System;
using System.Collections.Generic;
using Xunit;

namespace PanoramicData.SheetMagic.Test;

public class SaveSheetTests : Test
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
			new Dictionary<string, object?>
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
			_ = b.Should().NotBeNullOrEmpty();
			var firstItem = b[0];
			_ = firstItem.Properties.Keys.Should().Contain("a");
			_ = firstItem.Properties["a"].Should().Be("b");
		}
		finally
		{
			fileInfo.Delete();
		}
	}

	[Fact]
	public void SavingWithExtendedObjectContainingInt_Succeeds()
	{
		var a = new Extended<object>(
			new object(),
			new Dictionary<string, object?>
			{
				{ "a", 1 }
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
			_ = b.Should().NotBeNullOrEmpty();
			var firstItem = b[0];
			_ = firstItem.Properties.Keys.Should().Contain("a");
			_ = firstItem.Properties["a"].Should().Be(1);
		}
		finally
		{
			fileInfo.Delete();
		}
	}

	[Fact]
	public void SavingWithExtendedObjectContainingDateTime_Succeeds()
	{
		var dateTime = new DateTime(2000, 1, 2, 3, 4, 5);
		var a = new Extended<object>(
			new object(),
			new Dictionary<string, object?>
			{
				{ "a", dateTime }
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
			_ = b.Should().NotBeNullOrEmpty();
			var firstItem = b[0];
			_ = firstItem.Properties.Keys.Should().Contain("a");
			_ = firstItem.Properties["a"].Should().Be(dateTime);
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
			_ = cars.Should().NotBeNullOrEmpty();
			var firstCar = cars[0];
			_ = firstCar.Item.Should().NotBeNull();
			_ = carWeightKg.Should().Be(firstCar.Item!.WeightKg);
			_ = firstCar.Properties.Keys.Should().Contain(customPropertyName);
			_ = firstCar.Properties[customPropertyName].Should().Be(customPropertyValue);
		}
		finally
		{
			fileInfo.Delete();
		}
	}

	[Fact]
	public void TypesWithLists_Succeeds()
	{
		var fileInfo = GetXlsxTempFileInfo();

		var dealerships = new List<CarDealership>
		{
			new CarDealership
			{
				Name = "Slough",
				Cars = new List<Car?>(),
			},
			new CarDealership
			{
				Name = "Maidenhead",
				Cars = new List<Car?>
				{
					new Car
					{
						Name = "Ford Prefect",
						WeightKg = 1200
					},
					null,
				},
			},
			new CarDealership
			{
				Name = "Reading",
				Cars = new List<Car?>
				{
					new Car
					{
						Name = "Ford Prefect",
						WeightKg = 1200
					},
					new Car
					{
						Name = "Ford Focus",
						WeightKg = 1500
					},
				},
			},
		};

		try
		{
			using var s = new MagicSpreadsheet(fileInfo);
			s.AddSheet(dealerships);
			s.Save();
		}
		finally
		{
			fileInfo.Delete();
		}
	}

	[Fact]
	public void TablesWithSameDisplayNameShouldNotFail()
	{
		var a = new Extended<object>(
			new object(),
			new Dictionary<string, object?>
			{
				{ "a", "b" }
			}
		);
		var fileInfo = GetXlsxTempFileInfo();
		try
		{
			// Save
			using var s1 = new MagicSpreadsheet(fileInfo);
			s1.AddSheet(new List<Extended<object>> { a }, "Sheet A", new AddSheetOptions { TableOptions = new TableOptions { DisplayName = "Table1" } });
			s1.AddSheet(new List<Extended<object>> { a }, "Sheet B", new AddSheetOptions { TableOptions = new TableOptions { DisplayName = "Table1" } });
		}
		finally
		{
			fileInfo.Delete();
		}
	}

	[Fact]
	public void TablesWithAutoDisplayNameShouldSucceed()
	{
		var a = new Extended<object>(
			new object(),
			new Dictionary<string, object?>
			{
				{ "a", "b" }
			}
		);
		var fileInfo = GetXlsxTempFileInfo();

		try
		{
			// Save
			using var s1 = new MagicSpreadsheet(fileInfo);
			s1.AddSheet(new List<Extended<object>> { a }, "Sheet A");
			s1.AddSheet(new List<Extended<object>> { a }, "Sheet B");
			s1.Save();
		}
		finally
		{
			fileInfo.Delete();
		}
	}
}