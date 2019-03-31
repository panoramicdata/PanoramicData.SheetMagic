using PanoramicData.SheetMagic.Test.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Xunit;

namespace PanoramicData.SheetMagic.Test
{
	public class LoadSheet
	{
		[Fact]
		public void LoadSheet_WithBlankRow_Succeeds()
		{
			// Load the parent/child relationships
			List<ParentChildRelationship> parentChildRelationships;
			using (var magicSpreadsheet = new MagicSpreadsheet(GetSheetFileInfo("ParentChildWithBlankRows"), new Options { StopProcessingOnFirstEmptyRow = true }))
			{
				magicSpreadsheet.Load();
				parentChildRelationships = magicSpreadsheet.GetList<ParentChildRelationship>();
			}

			// Loaded
			Assert.Equal(3, parentChildRelationships.Count);
		}

		[Fact]
		public void LoadParentChild()
		{
			// Load the parent/child relationships
			using (var magicSpreadsheet = new MagicSpreadsheet(GetSheetFileInfo("ParentChild")))
			{
				magicSpreadsheet.Load();
				magicSpreadsheet.GetList<ParentChildRelationship>();
			}

			// Loaded
		}

		[Fact]
		public void LoadAbc()
		{
			// Load the parent/child relationships
			List<AbcThing> things;
			using (var magicSpreadsheet = new MagicSpreadsheet(GetSheetFileInfo("EnumTest")))
			{
				magicSpreadsheet.Load();
				things = magicSpreadsheet.GetList<AbcThing>();
			}

			// Loaded
			Assert.True(things.Count == 6);
			Assert.Equal(AbcEnum.B, things[1].AbcEnum);
		}

		[Fact]
		public void LoadParentChild_MissingColumns_ThrowsException()
		{
			// Load the parent/child relationships
			using (var magicSpreadsheet = new MagicSpreadsheet(GetSheetFileInfo("ParentChild")))
			{
				magicSpreadsheet.Load();
				Assert.ThrowsAny<Exception>(() => magicSpreadsheet.GetList<ExtendedParentChildRelationship>());
			}

			// Loaded
		}

		[Fact]
		public void LoadParentChild_MissingColumnsOptionSet_Succeeds()
		{
			// Load the parent/child relationships
			using (var magicSpreadsheet = new MagicSpreadsheet(GetSheetFileInfo("ParentChild"), new Options { IgnoreUnmappedProperties = true }))
			{
				magicSpreadsheet.Load();
				magicSpreadsheet.GetList<ExtendedParentChildRelationship>();
			}

			// Loaded
		}

		[Fact]
		public void WriteAndLoadBack()
		{
			var tempFileInfo = new FileInfo(Path.GetTempFileName());
			try
			{
				// Generate some data
				var funkyAnimals = GetFunkyAnimals();
				var cars = GetCars();

				using (var magicSpreadsheet = new MagicSpreadsheet(tempFileInfo))
				{
					magicSpreadsheet.AddSheet(funkyAnimals);
					magicSpreadsheet.AddSheet(cars);
					magicSpreadsheet.Save();
				}

				// Reload
				using (var magicSpreadsheet = new MagicSpreadsheet(tempFileInfo))
				{
					magicSpreadsheet.Load();
					var sheetNames = magicSpreadsheet.SheetNames;
					Assert.Contains("FunkyAnimals", sheetNames);
					Assert.Contains("Cars", sheetNames);

					var reloadedCars = magicSpreadsheet.GetList<Car>();
					Assert.Equal(cars.Count, reloadedCars.Count);

					var reloadedAnimals = magicSpreadsheet.GetList<FunkyAnimal>("FunkyAnimals");
					Assert.Equal(funkyAnimals.Count, reloadedAnimals.Count);
				}
			}
			finally
			{
				// Clean up
				tempFileInfo.Delete();
			}
		}

		[Fact]
		public void WriteAndLoadBackAsExtended()
		{
			var tempFileInfo = new FileInfo(Path.GetTempFileName());
			try
			{
				// Generate some data
				var funkyAnimals = GetFunkyAnimals();
				var cars = GetCars();

				using (var magicSpreadsheet = new MagicSpreadsheet(tempFileInfo))
				{
					magicSpreadsheet.AddSheet(funkyAnimals);
					magicSpreadsheet.AddSheet(cars);
					magicSpreadsheet.Save();
				}

				// Reload
				using (var magicSpreadsheet = new MagicSpreadsheet(tempFileInfo,
				new Options
				{
					LoadNullExtendedProperties = true
				}))
				{
					magicSpreadsheet.Load();
					var sheetNames = magicSpreadsheet.SheetNames;
					Assert.Contains("FunkyAnimals", sheetNames);
					Assert.Contains("Cars", sheetNames);

					var reloadedCars = magicSpreadsheet.GetList<Car>();
					Assert.Equal(cars.Count, reloadedCars.Count);

					var reloadedAnimals = magicSpreadsheet.GetExtendedList<SimpleAnimal>("FunkyAnimals");
					Assert.Equal(funkyAnimals.Count, reloadedAnimals.Count);
					// Make sure the extra fields are there in the additional items
					Assert.All(reloadedAnimals, extendedAnimal => Assert.NotNull(extendedAnimal.Item));
					Assert.All(reloadedAnimals, extendedAnimal => Assert.NotEqual(0, extendedAnimal.Item.Id));
					Assert.All(reloadedAnimals, extendedAnimal => Assert.NotNull(extendedAnimal.Item.Name));
					Assert.All(reloadedAnimals, extendedAnimal => Assert.NotNull(extendedAnimal.Properties));
					Assert.All(reloadedAnimals, extendedAnimal => Assert.NotEmpty(extendedAnimal.Properties));
					Assert.All(reloadedAnimals, extendedAnimal => Assert.NotNull(extendedAnimal.Properties.Select(p => p.Value)));
					Assert.All(reloadedAnimals, extendedAnimal => Assert.All(extendedAnimal.Properties.Where(p => p.Key != nameof(FunkyAnimal.Description)).Select(p => p.Value), Assert.NotNull));
				}
			}
			finally
			{
				// Clean up
				tempFileInfo.Delete();
			}
		}

		internal static List<FunkyAnimal> GetFunkyAnimals()
		{
			return new List<FunkyAnimal>
			{
				new FunkyAnimal {Id = 1, Name = "Pig", Leg_Count = 4, WeightKg = 100.5, Description = "Bald sheep"},
				new FunkyAnimal {Id = 2, Name = "Chicken", Leg_Count = 2, WeightKg = 0.5},
				new FunkyAnimal {Id = 3, Name = "Goat", Leg_Count = 4, WeightKg = 30}
			};
		}

		internal static List<Car> GetCars()
		{
			return new List<Car>
			{
				new Car {Id = 1, Name = "Ford Prefect", WheelCount = 4, WeightKg = 75},
				new Car {Id = 2, Name = "Ford! Focus!", WheelCount = 4, WeightKg = 2000}
			};
		}

		/// <summary>
		/// Tries to load bad sheets
		/// </summary>
		[Fact]
		public void LoadBadSheets()
		{
			var tempFileInfo = new FileInfo(Path.GetTempFileName());
			try
			{
				// Writes a sheet that has nothing to do with the attempt to read it
				var funkyAnimals = GetFunkyAnimals();

				using (var funkyAnimalMagicSpreadsheet = new MagicSpreadsheet(tempFileInfo))
				{
					funkyAnimalMagicSpreadsheet.AddSheet(funkyAnimals);
					funkyAnimalMagicSpreadsheet.Save();
				}

				using (var loadMagicSpreadsheet = new MagicSpreadsheet(tempFileInfo))
				{
					loadMagicSpreadsheet.Load();

					Assert.Throws<InvalidOperationException>(() => loadMagicSpreadsheet.GetList<Car>("Animals"));

					// Try to load FunkyAnimals into a list of Animals (should succeed)
					loadMagicSpreadsheet.GetList<Animal>("Animals");
				}
			}
			finally
			{
				tempFileInfo.Delete();
			}
		}

		[Fact]
		public void Load_Uae_Broken()
		{
			using (var sheet = new MagicSpreadsheet(GetSheetFileInfo("UAE_Broken"), new Options
			{
				EmptyRowInterpretedAsNull = true,
				StopProcessingOnFirstEmptyRow = false
			}))
			{
				sheet.Load();
				var deviceSpecifications = sheet.GetExtendedList<object>("Successes");
			}
		}

		[Fact]
		public void LoadSpreadsheet()
		{
			using (var sheet = new MagicSpreadsheet(GetSheetFileInfo("Bulk Import Template")))
			{
				sheet.Load();
				var deviceSpecifications = sheet.GetExtendedList<DeviceSpecification>();
				// do some sheet
				Assert.NotEmpty(deviceSpecifications);
				Assert.Single(deviceSpecifications);

				var device = deviceSpecifications[0];

				Assert.Equal("localhost", device.Item.HostName);
				Assert.Equal("DeviceDisplayName", device.Item.DeviceDisplayName);
				Assert.Equal("The device description", device.Item.DeviceDescription);
				Assert.Equal("Group/SubGroup1;Group/SubGroup2", device.Item.DeviceGroups);
				Assert.Equal("CollectorDescription", device.Item.PreferredCollector);
				Assert.True(device.Item.EnableAlerts);
				Assert.False(device.Item.EnableNetflow);
				Assert.Equal("", device.Item.NetflowCollector);
				Assert.Equal("http://www.logicmonitor.com/", device.Item.Link);

				// make sure there are 2 custom properties and are the values we're expecting
				Assert.Equal(2, device.Properties.Count);
				Assert.Equal("ValueA", device.Properties["Column A"]);
				Assert.Equal("ValueB", device.Properties["column.b"]);
			}
		}

		private static FileInfo GetSheetFileInfo(string worksheetName)
		{
			var location = typeof(LoadSheet).GetTypeInfo().Assembly.Location;
			var dirPath = Path.Combine(Path.GetDirectoryName(location), "../../../Sheets");
			return new FileInfo(Path.Combine(dirPath, $"{worksheetName}.xlsx"));
		}
	}
}