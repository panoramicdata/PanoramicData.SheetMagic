namespace PanoramicData.SheetMagic.Test;

public partial class SaveSheetTests : Test
{
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
				s1.AddSheet([car]);
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

		var dealerships = CreateDealershipsWithCars();

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

	private static List<CarDealershipWithCars> CreateDealershipsWithCars()
		=> new()
		{
			new() {
				Name = "Slough",
				Cars = [],
				Founded = new DateTime(2000, 1, 2, 3, 4, 5),
				UkRanking = 1,
				EmployeeCount = 10,
			},
			new() {
				Name = "Maidenhead",
				Cars =
				[
					new Car
					{
						Name = "Ford Prefect",
						WeightKg = 1200
					},
					null,
				],
				Founded = new DateTime(2000, 1, 2, 3, 4, 5),
				ClosureDate = new DateTime(2002, 1, 2, 3, 4, 5),
				UkRanking = 2,
				EmployeeCount = null,
			},
			new() {
				Name = "Reading",
				Cars =
				[
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
				],
				Founded = new DateTime(2000, 1, 2, 3, 4, 5),
				UkRanking = 3,
				EmployeeCount = 20,
			}
		};

	[Fact]
	public void TypesWithoutLists_Succeeds()
	{
		var fileInfo = GetXlsxTempFileInfo();

		var dealerships = new List<CarDealership>
		{
			new() {
				Name = "Slough",
				Founded = new DateTime(2000, 1, 2, 3, 4, 5),
				IsPrivatelyOwned = true,
				UkRanking = 1,
				EmployeeCount = 10,
			},
			new() {
				Name = "Maidenhead",
				Founded = new DateTime(2000, 1, 2, 3, 4, 5),
				ClosureDate = new DateTime(2002, 1, 2, 3, 4, 5),
				UkRanking = 2,
				EmployeeCount = null,
			},
			new() {
				Name = "Reading",
				Founded = new DateTime(2000, 1, 2, 3, 4, 5),
				UkRanking = 3,
				EmployeeCount = 20,
			},
		};

		try
		{
			using var s = new MagicSpreadsheet(fileInfo);
			s.AddSheet(dealerships);
			s.Save();

			using var s2 = new MagicSpreadsheet(fileInfo);
			s2.Load();
			var b = s2.GetList<CarDealership>();
			_ = b.Should().NotBeNullOrEmpty();
			var firstItem = b[0];
			firstItem.Should().NotBeNull();
			firstItem!.Name.Should().Be("Slough");
			firstItem.IsPrivatelyOwned.Should().BeTrue();
			firstItem.Founded.Should().Be(new DateTime(2000, 1, 2, 3, 4, 5));
			firstItem.UkRanking.Should().Be(1);
			firstItem.EmployeeCount.Should().Be(10);
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
			s1.AddSheet([a], "Sheet A", new AddSheetOptions { TableOptions = new TableOptions { DisplayName = "Table1" } });
			s1.AddSheet([a], "Sheet B", new AddSheetOptions { TableOptions = new TableOptions { DisplayName = "Table1" } });
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
			s1.AddSheet([a], "Sheet A");
			s1.AddSheet([a], "Sheet B");
			s1.Save();
		}
		finally
		{
			fileInfo.Delete();
		}
	}
}
