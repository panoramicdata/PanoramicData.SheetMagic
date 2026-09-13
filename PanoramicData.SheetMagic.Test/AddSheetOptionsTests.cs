using System.Drawing;

namespace PanoramicData.SheetMagic.Test;

public class AddSheetOptionsTests : Test
{
	[Fact]
	public void IncludeProperties_ListSpecified_CorrectProperties()
	{
		var fileInfo = GetXlsxTempFileInfo();

		try
		{
			var funkyAnimals = LoadSheetTests.GetFunkyAnimals();
			var options = new AddSheetOptions
			{
				TableOptions = new TableOptions
				{
					Name = "FunkyAnimals",
					DisplayName = "FunkyAnimals",
				},
				IncludeProperties =
				[
					nameof(SimpleAnimal.Id),
					nameof(SimpleAnimal.Name),
				]
			};
			using (var s = new MagicSpreadsheet(fileInfo))
			{
				s.AddSheet(funkyAnimals, "FunkyAnimals", options);
				s.Save();
			}
			// Reload the values back in and verify only the included properties exist

			using (var s = new MagicSpreadsheet(fileInfo))
			{
				s.Load();

				var reloadedAnimals = s.GetExtendedList<SimpleAnimal>("Sheet1");
				_ = reloadedAnimals.Should().HaveCount(funkyAnimals.Count);

				// Make sure there are no extra properties
				Assert.All(reloadedAnimals, static extendedAnimal => extendedAnimal.Properties.Should().BeEmpty());

				// Make sure items exist for every row
				Assert.All(reloadedAnimals, static extendedAnimal => extendedAnimal.Item.Should().NotBeNull());

				// Make sure that there are no "default" values we know are NOT in the test data
				Assert.All(reloadedAnimals, static extendedAnimal => Assert.NotEqual(0, extendedAnimal.Item!.Id));
				Assert.All(reloadedAnimals, static extendedAnimal => extendedAnimal.Item!.Name.Should().NotBe(string.Empty));
			}
		}
		finally
		{
			fileInfo.Delete();
		}
	}

	[Fact]
	public void ExcludeProperties_ListSpecified_CorrectProperties()
	{
		var fileInfo = GetXlsxTempFileInfo();

		try
		{
			var funkyAnimals = LoadSheetTests.GetFunkyAnimals();
			var options = new AddSheetOptions
			{
				TableOptions = new TableOptions
				{
					Name = "FunkyAnimals",
					DisplayName = "FunkyAnimals",
				},
				ExcludeProperties =
				[
					nameof(FunkyAnimal.Leg_Count),
					nameof(FunkyAnimal.WeightKg),
					nameof(FunkyAnimal.Description)
				]
			};
			using (var s = new MagicSpreadsheet(fileInfo))
			{
				s.AddSheet(funkyAnimals, "FunkyAnimals", options);
				s.Save();
			}
			// Reload the values back in and verify only the included properties exist

			using (var s = new MagicSpreadsheet(fileInfo))
			{
				s.Load();

				var reloadedAnimals = s.GetExtendedList<SimpleAnimal>("Sheet1");
				Assert.Equal(funkyAnimals.Count, reloadedAnimals.Count);
				// Make sure there are no extra properties

				var expectedProperties = new Dictionary<string, object?> {
					{ "Nicknames", "" },
					{ "Friends", "" }
					};

				Assert.All(reloadedAnimals, extendedAnimal => extendedAnimal.Properties.Should().BeEquivalentTo(expectedProperties));

				// Make sure items exist for every row
				Assert.All(reloadedAnimals, extendedAnimal => extendedAnimal.Item.Should().NotBeNull());

				// Make sure that there are no "default" values we know are NOT in the test data
				Assert.All(reloadedAnimals, extendedAnimal => Assert.NotEqual(0, extendedAnimal.Item!.Id));
				Assert.All(reloadedAnimals, static extendedAnimal => extendedAnimal.Item!.Name.Should().NotBe(string.Empty));
			}
		}
		finally
		{
			fileInfo.Delete();
		}
	}

	[Fact]
	public void AddSheetOptions_SheetWithExtendedPropertiesSorted_Succeeds()
	{
		var fileInfo = GetXlsxTempFileInfo();

		try
		{
			using var s = new MagicSpreadsheet(fileInfo);

			var sheetOptions = new AddSheetOptions
			{
				SortExtendedProperties = true
			};

			var animals = new Dictionary<string, object?> {
					  { "Type", "Hamster" },
					  { "Name", "Scruffy" }
				 };

			s.AddSheet(new List<Extended<object>>
				 {
					  new(new object(), animals)
				 }, "Animals", sheetOptions);
			s.Save();

			// Reload the values back in and verify only the included properties exist

			s.Load();

			var reloadedAnimals = s.GetExtendedList<object>("Animals");
			_ = reloadedAnimals.Should().NotBeNull();
			_ = reloadedAnimals.Should().HaveCountGreaterThan(0);
		}
		finally
		{
			fileInfo.Delete();
		}
	}

	[Fact]
	public void AddSheetOptions_SheetWithExtendedPropertiesUnsorted_Succeeds()
	{
		var fileInfo = GetXlsxTempFileInfo();

		try
		{
			using var s = new MagicSpreadsheet(fileInfo);

			var sheetOptions = new AddSheetOptions
			{
				SortExtendedProperties = false
			};

			var animals = new Dictionary<string, object?> {
					  { "Type", "Hamster" },
					  { "Name", "Scruffy" }
				 };

			s.AddSheet(new List<Extended<object>>
				 {
					  new(new object(), animals)
				 }, "Animals", sheetOptions);
			s.Save();

			// Reload the values back in and verify only the included properties exist

			s.Load();

			var reloadedAnimals = s.GetExtendedList<object>("Animals");
			_ = reloadedAnimals.Should().NotBeNull();
			_ = reloadedAnimals.Should().HaveCountGreaterThan(0);
		}
		finally
		{
			fileInfo.Delete();
		}
	}

	[Fact]
	public void AddSheetOptions_CustomTableStyle_Succeeds()
	{
		var fileInfo = GetXlsxTempFileInfo();

		try
		{
			using var s = new MagicSpreadsheet(fileInfo, new Options
			{
				TableStyles = [CreateCustomTableStyle(CustomTableStyleName)]
			});

			var sheetOptions = new AddSheetOptions
			{
				SortExtendedProperties = false,
				TableOptions = new TableOptions
				{
					CustomTableStyle = CustomTableStyleName
				}
			};

			s.AddSheet(CreateAnimals(), "Animals", sheetOptions);
			s.Save();
		}
		finally
		{
			fileInfo.Delete();
		}
	}

	private const string CustomTableStyleName = "My Table Style";

	private static CustomTableStyle CreateCustomTableStyle(string name)
		=> new()
		{
			Name = name,
			HeaderRowStyle = new TableRowStyle
			{
				BackgroundColor = Color.FromArgb(112, 48, 160),
				FontColor = Color.White,
				FontWeight = FontWeight.Bold
			},
			OddRowStyle = new TableRowStyle
			{
				BackgroundColor = Color.FromArgb(225, 204, 240),
			},
			EvenRowStyle = new TableRowStyle
			{
				BackgroundColor = Color.LightYellow,
			},
			WholeTableStyle = new TableRowStyle
			{
				InnerBorderColor = Color.Red,
				OuterBorderColor = Color.Blue
			},
		};

	private static List<Extended<object>> CreateAnimals()
		=>
		[
			CreateAnimal("Hamster", "Scruffy"),
			CreateAnimal("Dog", "Wuffy"),
			CreateAnimal("Fish", "Puffy"),
			CreateAnimal("Goat", "Gruffy")
		];

	private static Extended<object> CreateAnimal(string type, string name)
		=> new(new object(), new Dictionary<string, object?>
		{
			{ "Type", type },
			{ "Name", name }
		});
}
