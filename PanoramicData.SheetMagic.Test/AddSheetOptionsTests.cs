using FluentAssertions;
using PanoramicData.SheetMagic.Test.Models;
using System.Collections.Generic;
using System.Drawing;
using Xunit;

namespace PanoramicData.SheetMagic.Test
{
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
					IncludeProperties = new HashSet<string>
					{
						nameof(SimpleAnimal.Id),
						nameof(SimpleAnimal.Name),
					}
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
					reloadedAnimals.Count.Should().Be(funkyAnimals.Count);

					// Make sure there are no extra properties
					Assert.All(reloadedAnimals, extendedAnimal => Assert.Empty(extendedAnimal.Properties));

					// Make sure items exist for every row
					Assert.All(reloadedAnimals, extendedAnimal => Assert.NotNull(extendedAnimal.Item));

					// Make sure that there are no "default" values we know are NOT in the test data
					Assert.All(reloadedAnimals, extendedAnimal => Assert.NotEqual(0, extendedAnimal.Item!.Id));
					Assert.All(reloadedAnimals, extendedAnimal => Assert.NotEqual(string.Empty, extendedAnimal.Item!.Name));
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
					ExcludeProperties = new HashSet<string>
					{
						nameof(FunkyAnimal.Leg_Count),
						nameof(FunkyAnimal.WeightKg),
						nameof(FunkyAnimal.Description)
					}
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
					Assert.All(reloadedAnimals, extendedAnimal => Assert.Empty(extendedAnimal.Properties));

					// Make sure items exist for every row
					Assert.All(reloadedAnimals, extendedAnimal => Assert.NotNull(extendedAnimal.Item));

					// Make sure that there are no "default" values we know are NOT in the test data
					Assert.All(reloadedAnimals, extendedAnimal => Assert.NotEqual(0, extendedAnimal.Item!.Id));
					Assert.All(reloadedAnimals, extendedAnimal => Assert.NotEqual(string.Empty, extendedAnimal.Item!.Name));
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
						  new Extended<object>(new object(), animals)
					 }, "Animals", sheetOptions);
				s.Save();

				// Reload the values back in and verify only the included properties exist

				s.Load();

				// TODO
				var reloadedAnimals = s.GetExtendedList<object>("Animals");
				//Assert.Equal(reloadedAnimals.[0].Key, "Name");
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
						  new Extended<object>(new object(), animals)
					 }, "Animals", sheetOptions);
				s.Save();

				// Reload the values back in and verify only the included properties exist

				s.Load();

				// TODO
				var reloadedAnimals = s.GetExtendedList<object>("Animals");
				//Assert.Equal(reloadedAnimals.[0].Key, "Type");
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
					TableStyles = new List<CustomTableStyle>
					{
						new CustomTableStyle
						{
							Name = "My Table Style",
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
						}
					}
				});

				var sheetOptions = new AddSheetOptions
				{
					SortExtendedProperties = false,
					TableOptions = new TableOptions
					{
						CustomTableStyle = "My Table Style"
					}
				};

				var scruffy = new Dictionary<string, object?> {
						  { "Type", "Hamster" },
						  { "Name", "Scruffy" }
					 };
				var wuffy = new Dictionary<string, object?> {
						  { "Type", "Dog" },
						  { "Name", "Wuffy" }
					 };
				var puffy = new Dictionary<string, object?> {
						  { "Type", "Fish" },
						  { "Name", "Puffy" }
					 };
				var gruffy = new Dictionary<string, object?> {
						  { "Type", "Goat" },
						  { "Name", "Gruffy" }
					 };

				s.AddSheet(new List<Extended<object>>
					 {
						  new Extended<object>(new object(), scruffy),
						  new Extended<object>(new object(), wuffy),
						  new Extended<object>(new object(), puffy),
						  new Extended<object>(new object(), gruffy)
					 }, "Animals", sheetOptions);
				s.Save();
			}
			finally
			{
				fileInfo.Delete();
			}
		}
	}
}
