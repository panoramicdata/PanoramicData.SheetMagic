using FluentAssertions;
using PanoramicData.SheetMagic.Test.Models;
using System.Collections.Generic;
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
				var funkyAnimals = LoadSheet.GetFunkyAnimals();
				var options = new AddSheetOptions
				{
					TableOptions = new TableOptions
					{
						Name = "FunkyAnimals",
						DisplayName = "Funky Animals",
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
				var funkyAnimals = LoadSheet.GetFunkyAnimals();
				var options = new AddSheetOptions
				{
					TableOptions = new TableOptions
					{
						Name = "FunkyAnimals",
						DisplayName = "Funky Animals",
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
	}
}
