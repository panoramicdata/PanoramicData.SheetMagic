using PanoramicData.SheetMagic.Test.Models;
using System.Collections.Generic;
using System.IO;
using Xunit;

namespace PanoramicData.SheetMagic.Test
{
	public class AddSheetOptionsTests
	{
		[Fact]
		public void IncludeProperties_ListSpecified_CorrectProperties()
		{
			var fileInfo = new FileInfo(Path.GetTempFileName());

			try
			{
				var funkyAnimals = LoadSheet.GetFunkyAnimals();
				var options = new AddSheetOptions
				{
					IncludeProperties = new HashSet<string>
					{
						nameof(SimpleAnimal.Id),
						nameof(SimpleAnimal.Name),
					}
				};
				using (var s = new MagicSpreadsheet(fileInfo))
				{
					s.AddSheet(funkyAnimals, "Sheet1", options);
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
					Assert.All(reloadedAnimals, extendedAnimal => Assert.NotEqual(0, extendedAnimal.Item.Id));
					Assert.All(reloadedAnimals, extendedAnimal => Assert.NotEqual(string.Empty, extendedAnimal.Item.Name));
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
			var fileInfo = new FileInfo(Path.GetTempFileName());

			try
			{
				var funkyAnimals = LoadSheet.GetFunkyAnimals();
				var options = new AddSheetOptions
				{
					ExcludeProperties = new HashSet<string>
					{
						nameof(FunkyAnimal.Leg_Count),
						nameof(FunkyAnimal.WeightKg),
						nameof(FunkyAnimal.Description)
					}
				};
				using (var s = new MagicSpreadsheet(fileInfo))
				{
					s.AddSheet(funkyAnimals, "Sheet1", options);
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
					Assert.All(reloadedAnimals, extendedAnimal => Assert.NotEqual(0, extendedAnimal.Item.Id));
					Assert.All(reloadedAnimals, extendedAnimal => Assert.NotEqual(string.Empty, extendedAnimal.Item.Name));
				}
			}
			finally
			{
				fileInfo.Delete();
			}
		}
	}
}
