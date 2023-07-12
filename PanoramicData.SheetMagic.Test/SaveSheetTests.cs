﻿using FluentAssertions;
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
		var theString = "b";

		Check(theString);
	}

	[Theory]
	[InlineData(int.MinValue)]
	[InlineData(-1)]
	[InlineData(0)]
	[InlineData(1)]
	[InlineData(int.MaxValue)]
	public void SavingWithExtendedObjectContainingInt32_Succeeds(long value)
		=> Check(value);

	[Theory]
	[InlineData((uint)0)]
	[InlineData((uint)1)]
	[InlineData(uint.MaxValue)]
	public void SavingWithExtendedObjectContainingUInt32_Succeeds(uint value)
		=> Check(value);

	[Theory]
	[InlineData(long.MinValue)]
	[InlineData((long)-1)]
	[InlineData((long)0)]
	[InlineData((long)1)]
	[InlineData(long.MaxValue)]
	public void SavingWithExtendedObjectContainingInt64_Succeeds(long value)
		=> Check(value);

	[Theory]
	[InlineData((ulong)0)]
	[InlineData((ulong)1)]
	[InlineData(ulong.MaxValue)]
	public void SavingWithExtendedObjectContainingUInt64_Succeeds(ulong value)
		=> Check(value);

	[Theory]
	[InlineData(short.MinValue)]
	[InlineData((short)-1)]
	[InlineData((short)0)]
	[InlineData((short)1)]
	[InlineData(short.MaxValue)]
	public void SavingWithExtendedObjectContainingInt16_Succeeds(short value)
		=> Check(value);


	[Theory]
	[InlineData((ushort)0)]
	[InlineData((ushort)1)]
	[InlineData(ushort.MaxValue)]
	public void SavingWithExtendedObjectContainingUInt16_Succeeds(ushort value)
		=> Check(value);

	[Theory]
	[InlineData(true)]
	[InlineData(false)]
	public void SavingWithExtendedObjectContainingBoolean_Succeeds(bool inputBool)
		=> Check(inputBool);

	[Theory]
	[InlineData(true)]
	[InlineData(false)]
	[InlineData(null)]
	public void SavingWithExtendedObjectContainingNullableBoolean_Succeeds(bool? inputBool)
		=> Check(inputBool);

	[Theory]
	[InlineData(double.MinValue)]
	[InlineData(double.MaxValue)]
	[InlineData(double.NegativeInfinity)]
	[InlineData(double.PositiveInfinity)]
	[InlineData(double.NaN)]
	[InlineData(12.3)]
	[InlineData(-1.0)]
	public void SavingWithExtendedObjectContainingDouble_Succeeds(double value)
		=> Check(value);

	[Fact]
	public void SavingWithExtendedObjectContainingDateTime_Succeeds()
		=> Check(new DateTime(2000, 1, 2, 3, 4, 5));

	[Fact]
	public void SavingWithExtendedObjectContainingNullableDateTime_Succeeds()
		=> Check((DateTime?)new DateTime(2000, 1, 2, 3, 4, 5));

	[Fact]
	public void SavingWithExtendedObjectContainingDateTimeOffset_Succeeds()
		=> Check(new DateTimeOffset(2000, 1, 2, 3, 4, 5, TimeSpan.Zero));

	[Fact]
	public void SavingWithExtendedObjectContainingNullableDateTimeOffset_Succeeds()
		=> Check((DateTimeOffset?)new DateTimeOffset(2000, 1, 2, 3, 4, 5, TimeSpan.Zero));

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

	private static void Check<T>(T theValue)
	{
		var a = new Extended<object>(
			new object(),
			new Dictionary<string, object?>
			{
				{ "a", theValue }
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
			_ = firstItem.Properties["a"].Should().Be(theValue);
		}
		finally
		{
			fileInfo.Delete();
		}
	}
}