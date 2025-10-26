using AwesomeAssertions;
using PanoramicData.SheetMagic.Test.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace PanoramicData.SheetMagic.Test;

public class LoadSheetTests : Test
{
	[Fact]
	public void LoadSheet_WithBlankRow_Succeeds()
	{
		// Load the parent/child relationships
		List<ParentChildRelationship?> parentChildRelationships;
		using (var magicSpreadsheet = new MagicSpreadsheet(GetSheetFileInfo("ParentChildWithBlankRows"), new Options { StopProcessingOnFirstEmptyRow = true }))
		{
			magicSpreadsheet.Load();
			parentChildRelationships = magicSpreadsheet.GetList<ParentChildRelationship>();
		}

		// Loaded
		_ = parentChildRelationships.Should().HaveCount(3);
	}

	[Fact]
	public void LoadSheet_WithBinaryValues_Succeeds()
	{
		// Load the parent/child relationships
		using var magicSpreadsheet = new MagicSpreadsheet(GetSheetFileInfo("LMREP-7413"), new Options { StopProcessingOnFirstEmptyRow = true });
		magicSpreadsheet.Load();
		var values = magicSpreadsheet.GetExtendedList<object>();
		_ = ((bool?)values[0].Properties["IncludeSection2"]).Should().BeTrue();
	}

	[Fact]
	public void LoadParentChild()
	{
		// Load the parent/child relationships
		using var magicSpreadsheet = new MagicSpreadsheet(GetSheetFileInfo("ParentChild"));
		magicSpreadsheet.Load();
		_ = magicSpreadsheet.GetList<ParentChildRelationship>();

		// Loaded
	}

	[Fact]
	public void LoadAbc()
	{
		// Load the parent/child relationships
		List<AbcThing?> things;
		using (var magicSpreadsheet = new MagicSpreadsheet(GetSheetFileInfo("EnumTest")))
		{
			magicSpreadsheet.Load();
			things = magicSpreadsheet.GetList<AbcThing>();
		}

		// Loaded
		_ = things.Should().HaveCount(6);
		_ = (things[1]?.AbcEnum.Should().Be(AbcEnum.B));
	}

	[Fact]
	public void LoadParentChild_MissingColumns_ThrowsException()
	{
		// Load the parent/child relationships
		using var magicSpreadsheet = new MagicSpreadsheet(GetSheetFileInfo("ParentChild"));
		magicSpreadsheet.Load();
		_ = Assert.ThrowsAny<Exception>(() => magicSpreadsheet.GetList<ExtendedParentChildRelationship>());
		// Loaded
	}

	[Fact]
	public void LoadParentChild_MissingColumnsOptionSet_Succeeds()
	{
		// Load the parent/child relationships
		using var magicSpreadsheet = new MagicSpreadsheet(GetSheetFileInfo("ParentChild"), new Options { IgnoreUnmappedProperties = true });
		magicSpreadsheet.Load();
		_ = magicSpreadsheet.GetList<ExtendedParentChildRelationship>();
		// Loaded
	}

	[Fact]
	public void WriteAndLoadBack()
	{
		var tempFileInfo = GetXlsxTempFileInfo();
		try
		{
			// Generate some data
			var funkyAnimals = GetFunkyAnimals();
			var cars = GetCars();

			using (var magicSpreadsheet = new MagicSpreadsheet(tempFileInfo))
			{
				// Exclude the Friends property as nested objects are not supported for reading from Excel
				var addSheetOptions = new AddSheetOptions
				{
					ExcludeProperties = ["Friends"]
				};
				magicSpreadsheet.AddSheet(funkyAnimals, null, addSheetOptions);
				magicSpreadsheet.AddSheet(cars);
				magicSpreadsheet.Save();
			}

			// Reload - need to ignore unmapped properties since we excluded Friends when saving
			using (var magicSpreadsheet = new MagicSpreadsheet(tempFileInfo, new Options { IgnoreUnmappedProperties = true }))
			{
				magicSpreadsheet.Load();
				var sheetNames = magicSpreadsheet.SheetNames;
				_ = sheetNames.Should().Contain("FunkyAnimals");
				_ = sheetNames.Should().Contain("Cars");

				var reloadedCars = magicSpreadsheet.GetList<Car>();
				_ = reloadedCars.Should().HaveCount(cars.Count);

				var reloadedAnimals = magicSpreadsheet.GetList<FunkyAnimal>("FunkyAnimals");
				_ = reloadedAnimals.Should().HaveCount(funkyAnimals.Count);
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
		var tempFileInfo = GetXlsxTempFileInfo();
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
				_ = sheetNames.Should().Contain("FunkyAnimals");
				_ = sheetNames.Should().Contain("Cars");

				var reloadedCars = magicSpreadsheet.GetList<Car>();
				_ = reloadedCars.Should().HaveCount(cars.Count);

				var reloadedAnimals = magicSpreadsheet.GetExtendedList<SimpleAnimal>("FunkyAnimals");
				_ = reloadedAnimals.Should().HaveCount(funkyAnimals.Count);
				// Make sure the extra fields are there in the additional items
				Assert.All(reloadedAnimals, static extendedAnimal => extendedAnimal.Item.Should().NotBeNull());
				Assert.All(reloadedAnimals, static extendedAnimal => Assert.NotEqual(0, extendedAnimal.Item!.Id));
				Assert.All(reloadedAnimals, static extendedAnimal => extendedAnimal.Item!.Name.Should().NotBeNull());
				Assert.All(reloadedAnimals, static extendedAnimal => extendedAnimal.Properties.Should().NotBeNull());
				Assert.All(reloadedAnimals, static extendedAnimal => Assert.NotEmpty(extendedAnimal.Properties));
				Assert.All(reloadedAnimals, static extendedAnimal => extendedAnimal.Properties.Select(static p => p.Value).Should().NotBeNull());
				Assert.All(reloadedAnimals, static extendedAnimal => Assert.All(extendedAnimal.Properties.Where(static p => p.Key != nameof(FunkyAnimal.Description)).Select(static p => p.Value), Assert.NotNull));
			}
		}
		finally
		{
			// Clean up
			tempFileInfo.Delete();
		}
	}

	internal static List<FunkyAnimal> GetFunkyAnimals()
		=>
		[
			new FunkyAnimal { Id = 1, Name = "Pig", Leg_Count = 4, WeightKg = 100.5, Description = "Bald sheep" },
			new FunkyAnimal { Id = 2, Name = "Chicken", Leg_Count = 2, WeightKg = 0.5 },
			new FunkyAnimal { Id = 3, Name = "Goat", Leg_Count = 4, WeightKg = 30 }
		];

	internal static List<Car> GetCars()
		=>
		[
			new Car { Id = 1, Name = "Ford Prefect", WheelCount = 4, WeightKg = 75 },
			new Car { Id = 2, Name = "Ford! Focus!", WheelCount = 4, WeightKg = 2000 }
		];

	/// <summary>
	/// Tries to load bad sheets
	/// </summary>
	[Fact]
	public void LoadBadSheets()
	{
		var tempFileInfo = GetXlsxTempFileInfo();
		try
		{
			// Writes a sheet that has nothing to do with the attempt to read it
			var funkyAnimals = GetFunkyAnimals();

			using (var funkyAnimalMagicSpreadsheet = new MagicSpreadsheet(tempFileInfo))
			{
				funkyAnimalMagicSpreadsheet.AddSheet(funkyAnimals);
				funkyAnimalMagicSpreadsheet.Save();
			}

			using var loadMagicSpreadsheet = new MagicSpreadsheet(tempFileInfo);
			loadMagicSpreadsheet.Load();

			_ = Assert.Throws<InvalidOperationException>(() => loadMagicSpreadsheet.GetList<Car>("Animals"));

			// Try to load FunkyAnimals into a list of Animals (should succeed)
			_ = loadMagicSpreadsheet.GetList<Animal>("Animals");
		}
		finally
		{
			tempFileInfo.Delete();
		}
	}

	[Fact]
	public void Load_Uae_Broken()
	{
		using var sheet = new MagicSpreadsheet(GetSheetFileInfo("UAE_Broken"), new Options
		{
			EmptyRowInterpretedAsNull = true,
			StopProcessingOnFirstEmptyRow = false
		});
		sheet.Load();
	}

	[Fact]
	public void Load_StarkIndustries_ShouldLoad()
	{
		using var sheet = new MagicSpreadsheet(GetSheetFileInfo("Stark Industries"),
			new Options
			{
				StopProcessingOnFirstEmptyRow = true
			}
		);
		sheet.Load();

		var sites = sheet.GetExtendedList<ImportedSite>("Sites");
		var siteItem = sites[1].Item;
		_ = siteItem.Should().NotBeNull();
		_ = siteItem!.FloorHeightFeet.Should().NotBe(0);
		_ = sites.Should().NotBeEmpty();

		var devices = sheet.GetExtendedList<ImportedDevice>("Devices");
		_ = devices.Should().NotBeEmpty();
	}

	[Fact]
	public void Load_SitesAndDevices_ShouldLoad()
	{
		using var sheet = new MagicSpreadsheet(GetSheetFileInfo("SitesAndDevices"),
			new Options
			{
				StopProcessingOnFirstEmptyRow = true
			}
		);
		sheet.Load();

		var sites = sheet.GetExtendedList<ImportedSite>("Sites");
		_ = sites.Should().NotBeEmpty();

		var devices = sheet.GetExtendedList<ImportedDevice>("Devices");
		_ = devices.Should().NotBeEmpty();
	}

	[Fact]
	public void Load_SitesAndNoDevices_TryToLoadFromMissingWorkSheetShouldThrowException()
	{
		using var sheet = new MagicSpreadsheet(GetSheetFileInfo("SitesAndDevices"),
			new Options
			{
				// StopProcessingOnFirstEmptyRow = true
				EmptyRowInterpretedAsNull = false,
				StopProcessingOnFirstEmptyRow = false
			}
		);
		sheet.Load();

		_ = Assert.ThrowsAny<Exception>(() => sheet.GetExtendedList<ImportedDevice>("XXX"));
	}

	[Fact]
	public void LoadSpreadsheet()
	{
		using var sheet = new MagicSpreadsheet(GetSheetFileInfo("Bulk Import Template"));
		sheet.Load();
		var deviceSpecifications = sheet.GetExtendedList<DeviceSpecification>();
		// do some sheet
		Assert.NotEmpty(deviceSpecifications);
		_ = Assert.Single(deviceSpecifications);

		var device = deviceSpecifications[0];
		_ = device.Item.Should().NotBeNull();
		_ = device.Item!.HostName.Should().Be("localhost");
		_ = device.Item.DeviceDisplayName.Should().Be("DeviceDisplayName");
		_ = device.Item.DeviceDescription.Should().Be("The device description");
		_ = device.Item.DeviceGroups.Should().Be("Group/SubGroup1;Group/SubGroup2");
		_ = device.Item.PreferredCollector.Should().Be("CollectorDescription");
		_ = device.Item.EnableAlerts.Should().Be(true);
		_ = device.Item.EnableNetflow.Should().Be(false);
		_ = device.Item.NetflowCollector.Should().Be(string.Empty);
		_ = device.Item.Link.Should().Be("http://www.logicmonitor.com/");

		// make sure there are 2 custom properties and are the values we're expecting
		_ = device.Properties.Should().HaveCount(2);
		_ = device.Properties["Column A"].Should().Be("ValueA");
		_ = device.Properties["column.b"].Should().Be("ValueB");
	}
}