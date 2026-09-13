using System.Text.Json;
using System.Text.Json.Nodes;

namespace PanoramicData.SheetMagic.Test;

public class AddSheetTests : Test
{
	[Theory]
	[InlineData("12345678901234567890123456789012")]
	[InlineData("abcdefghijklmnopqrstuvwxyz123456")]
	public void AddSheet_SheetNameTooLong_Fails(string badSheetName)
	{
		var fileInfo = GetXlsxTempFileInfo();
		var items = new List<SimpleAnimal> { new() { Id = 1, Name = "Alligator" } };

		try
		{
			using var s = new MagicSpreadsheet(fileInfo);
			Action act = () => s.AddSheet(items, badSheetName);
			act.Should().ThrowExactly<ArgumentException>();
		}
		finally
		{
			fileInfo.Delete();
		}
	}

	[Fact]
	public void AddSheet_JsonObjects_WithExtendedObject_Succeeds()
	{
		var fileInfo = GetXlsxTempFileInfo();

		try
		{
			using var s = new MagicSpreadsheet(fileInfo);
			var jsonObjectList = CreateAnimalJsonObjects();

			// Convert JsonObjects to Extended<object>
			var extendedList = new List<Extended<object>>();
			foreach (var jsonObject in jsonObjectList)
			{
				var extended = new Extended<object>(new(), jsonObject.Deserialize<Dictionary<string, object?>>() ?? throw new ArgumentException("Could not convert JsonObject to dictionary"));
				extendedList.Add(extended);
			}

			s.AddSheet(
				extendedList
			);

			s.Save();
		}
		finally
		{
			fileInfo.Delete();
		}
	}

	[Fact(Skip = "JsonObject support is not yet implemented. Use Extended<JsonObject> instead.")]
	public void AddSheet_JsonObjects_Succeeds()
	{
		var fileInfo = GetXlsxTempFileInfo();

		try
		{
			using var s = new MagicSpreadsheet(fileInfo);
			var jsonObjectList = CreateAnimalJsonObjects();

			s.AddSheet(
				jsonObjectList
			);

			s.Save();
		}
		finally
		{
			fileInfo.Delete();
		}
	}

	private static List<JsonObject> CreateAnimalJsonObjects() =>
		[
			ToJsonObject(new SimpleAnimal { Id = 1, Name = "alligator" }),
			ToJsonObject(new SimpleAnimal { Id = 2, Name = "bee" })
		];

	private static JsonObject ToJsonObject(SimpleAnimal animal)
		=> JsonSerializer.SerializeToNode(animal)?.AsObject()
			?? throw new ArgumentException("Could not convert the animal to a JsonObject", nameof(animal));

	[Fact]
	public void AddSheet_SheetNameAlreadyExists_Fails()
	{
		var fileInfo = GetXlsxTempFileInfo();

		try
		{
			using var s = new MagicSpreadsheet(fileInfo);
			var items = new List<SimpleAnimal> { new() { Id = 1, Name = "Alligator" } };
			s.AddSheet(items, "Sheet1");
			Action act = () => s.AddSheet(items, "Sheet1");
			act.Should().ThrowExactly<ArgumentException>();
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
					  new(new object(), new Dictionary<string, object?> {
						  { "Id", 10 },
						  { "My Name", "Ryan" }
					  })
				 }, "Subscriptions", sheetOptions);
			s.Save();
		}
		finally
		{
			fileInfo.Delete();
		}
	}

	[Fact]
	public void AddSheet_SheetWithStyle_Succeeds()
	{
		var fileInfo = GetXlsxTempFileInfo();

		try
		{
			using var s = new MagicSpreadsheet(fileInfo);

			var sheetOptions = new AddSheetOptions
			{
				TableOptions = new TableOptions
				{
					XlsxTableStyle = XlsxTableStyle.TableStyleDark1
				}
			};

			s.AddSheet(new List<FunkyAnimal>
				 {
					new() { Id = 0, Name = "Old Woman", WeightKg = 60, Leg_Count = 2},
					new() { Id = 1, Name = "Horse", WeightKg = 200, Leg_Count = 4},
					new() { Id = 2, Name = "Cow", WeightKg = 100, Leg_Count = 4},
					new() { Id = 3, Name = "Dog", WeightKg = 50, Leg_Count = 4},
					new() { Id = 4, Name = "Cat", WeightKg = 25, Leg_Count = 4},
					new() { Id = 5, Name = "Mouse", WeightKg = 0.1, Leg_Count = 4},
					new() { Id = 7, Name = "Spider", WeightKg = 0.01, Leg_Count = 8},
					new() { Id = 8, Name = "Fly", WeightKg = 0.001, Leg_Count = 6}
				 }, "Animals", sheetOptions);
			s.Save();
		}
		finally
		{
			fileInfo.Delete();
		}
	}

	[Fact]
	public void AddSheet_MultipleSheetsWithStyle_Succeeds()
	{
		var fileInfo = GetXlsxTempFileInfo();

		try
		{
			using var s = new MagicSpreadsheet(fileInfo);

			var funkyAnimals = new List<FunkyAnimal>
			{
				new() { Id = 0, Name = "Old Woman", WeightKg = 60, Leg_Count = 2},
				new() { Id = 1, Name = "Horse", WeightKg = 200, Leg_Count = 4},
				new() { Id = 2, Name = "Cow", WeightKg = 100, Leg_Count = 4},
				new() { Id = 3, Name = "Dog", WeightKg = 50, Leg_Count = 4},
				new() { Id = 4, Name = "Cat", WeightKg = 25, Leg_Count = 4},
				new() { Id = 5, Name = "Mouse", WeightKg = 0.1, Leg_Count = 4},
				new() { Id = 7, Name = "Spider", WeightKg = 0.01, Leg_Count = 8},
				new() { Id = 8, Name = "Fly", WeightKg = 0.001, Leg_Count = 6}
			};

			var sheetOptions = new AddSheetOptions
			{
				TableOptions = new TableOptions
				{
					Name = "Table 1",
					DisplayName = "Table1",
					XlsxTableStyle = XlsxTableStyle.TableStyleDark1
				}
			};
			s.AddSheet(funkyAnimals, "Animals", sheetOptions);
			sheetOptions = new AddSheetOptions
			{
				TableOptions = new TableOptions
				{
					Name = "Table 2",
					DisplayName = "Table2",
					XlsxTableStyle = XlsxTableStyle.TableStyleDark2
				}
			};
			s.AddSheet(funkyAnimals, "Animals 2", sheetOptions);
			s.Save();
		}
		finally
		{
			fileInfo.Delete();
		}
	}

	[Fact]
	public void AddSheet_EnumerableProperty_Succeeds()
	{
		var fileInfo = GetXlsxTempFileInfo();

		try
		{
			using var s = new MagicSpreadsheet(fileInfo);

			var funkyAnimals = CreateFunkyAnimals();

			s.AddSheet(funkyAnimals, "Animals", CreateExpandedEnumerableSheetOptions());
			s.AddSheet(funkyAnimals, "Animals 2", CreateSecondTableSheetOptions());
			s.Save();
		}
		finally
		{
			fileInfo.Delete();
		}
	}

	[Fact]
	public void AddSheet_EnumerablePropertyWithExtended_Succeeds()
	{
		var fileInfo = GetXlsxTempFileInfo();

		try
		{
			using var s = new MagicSpreadsheet(fileInfo);
			var funkyAnimals = CreateExtendedFunkyAnimals();

			s.AddSheet(funkyAnimals, "Animals", CreateExpandedEnumerableSheetOptions());
			s.AddSheet(funkyAnimals, "Animals 2", CreateSecondTableSheetOptions());
			s.Save();
		}
		finally
		{
			fileInfo.Delete();
		}
	}

	private static List<FunkyAnimal> CreateFunkyAnimals()
		=>
		[
			new() {
				Id = 0,
				Name = "Old Woman",
				WeightKg = 60,
				Leg_Count = 2,
				Nicknames = ["Woo", "Yay"],
				Friends = [
					new() { Id = 10, Name="Houpla" },
					new() { Id = 11, Name = "Bedoink" }
				]
			},
			new() { Id = 1, Name = "Horse", WeightKg = 200, Leg_Count = 4, Nicknames = ["Bert", "Ernie"]},
			new() { Id = 2, Name = "Cow", WeightKg = 100, Leg_Count = 4},
			new() { Id = 3, Name = "Dog", WeightKg = 50, Leg_Count = 4},
			new() { Id = 4, Name = "Cat", WeightKg = 25, Leg_Count = 4},
			new() { Id = 5, Name = "Mouse", WeightKg = 0.1, Leg_Count = 4},
			new() { Id = 7, Name = "Spider", WeightKg = 0.01, Leg_Count = 8},
			new() { Id = 8, Name = "Fly", WeightKg = 0.001, Leg_Count = 6}
		];

	/// <summary>
	/// The same animals as <see cref="CreateFunkyAnimals"/>, where only the first carries an
	/// extended property, so that a sheet mixing the two is exercised.
	/// </summary>
	private static List<Extended<FunkyAnimal>> CreateExtendedFunkyAnimals()
	{
		var animals = CreateFunkyAnimals();
		var emptyDictionary = new Dictionary<string, object?>();

		return
		[
			new(animals[0], new Dictionary<string, object?> { { "Extended", "Extended" } }),
			.. animals.Skip(1).Select(animal => new Extended<FunkyAnimal>(animal, emptyDictionary))
		];
	}

	private static AddSheetOptions CreateExpandedEnumerableSheetOptions()
		=> new()
		{
			EnumerableCellOptions = new()
			{
				Expand = true,
				CellDelimiter = ", "
			}
		};

	private static AddSheetOptions CreateSecondTableSheetOptions()
		=> new()
		{
			TableOptions = new TableOptions
			{
				Name = "Table 2",
				DisplayName = "Table2",
				XlsxTableStyle = XlsxTableStyle.TableStyleDark2
			}
		};
}
