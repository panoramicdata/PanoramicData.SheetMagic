using FluentAssertions;
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
		parentChildRelationships.Count.Should().Be(3);
	}

	[Fact]
	public void LoadSheet_WithBinaryValues_Succeeds()
	{
		// Load the parent/child relationships
		using var magicSpreadsheet = new MagicSpreadsheet(GetSheetFileInfo("LMREP-7413"), new Options { StopProcessingOnFirstEmptyRow = true });
		magicSpreadsheet.Load();
		var values = magicSpreadsheet.GetExtendedList<object>();
		((bool?)values[0].Properties["IncludeSection2"]).Should().BeTrue();
	}

	[Fact]
	public void LoadParentChild()
	{
		// Load the parent/child relationships
		using var magicSpreadsheet = new MagicSpreadsheet(GetSheetFileInfo("ParentChild"));
		magicSpreadsheet.Load();
		magicSpreadsheet.GetList<ParentChildRelationship>();

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
		things.Count.Should().Be(6);
		things[1]?.AbcEnum.Should().Be(AbcEnum.B);
	}

	[Fact]
	public void LoadParentChild_MissingColumns_ThrowsException()
	{
		// Load the parent/child relationships
		using var magicSpreadsheet = new MagicSpreadsheet(GetSheetFileInfo("ParentChild"));
		magicSpreadsheet.Load();
		Assert.ThrowsAny<Exception>(() => magicSpreadsheet.GetList<ExtendedParentChildRelationship>());
		// Loaded
	}

	[Fact]
	public void LoadParentChild_MissingColumnsOptionSet_Succeeds()
	{
		// Load the parent/child relationships
		using var magicSpreadsheet = new MagicSpreadsheet(GetSheetFileInfo("ParentChild"), new Options { IgnoreUnmappedProperties = true });
		magicSpreadsheet.Load();
		magicSpreadsheet.GetList<ExtendedParentChildRelationship>();
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
				magicSpreadsheet.AddSheet(funkyAnimals);
				magicSpreadsheet.AddSheet(cars);
				magicSpreadsheet.Save();
			}

			// Reload
			using (var magicSpreadsheet = new MagicSpreadsheet(tempFileInfo))
			{
				magicSpreadsheet.Load();
				var sheetNames = magicSpreadsheet.SheetNames;
				sheetNames.Should().Contain("FunkyAnimals");
				sheetNames.Should().Contain("Cars");

				var reloadedCars = magicSpreadsheet.GetList<Car>();
				reloadedCars.Count.Should().Be(cars.Count);

				var reloadedAnimals = magicSpreadsheet.GetList<FunkyAnimal>("FunkyAnimals");
				reloadedAnimals.Count.Should().Be(funkyAnimals.Count);
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
				sheetNames.Should().Contain("FunkyAnimals");
				sheetNames.Should().Contain("Cars");

				var reloadedCars = magicSpreadsheet.GetList<Car>();
				reloadedCars.Count.Should().Be(cars.Count);

				var reloadedAnimals = magicSpreadsheet.GetExtendedList<SimpleAnimal>("FunkyAnimals");
				reloadedAnimals.Count.Should().Be(funkyAnimals.Count);
				// Make sure the extra fields are there in the additional items
				Assert.All(reloadedAnimals, extendedAnimal => Assert.NotNull(extendedAnimal.Item));
				Assert.All(reloadedAnimals, extendedAnimal => Assert.NotEqual(0, extendedAnimal.Item!.Id));
				Assert.All(reloadedAnimals, extendedAnimal => Assert.NotNull(extendedAnimal.Item!.Name));
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
		=> new()
		{
			new FunkyAnimal { Id = 1, Name = "Pig", Leg_Count = 4, WeightKg = 100.5, Description = "Bald sheep" },
			new FunkyAnimal { Id = 2, Name = "Chicken", Leg_Count = 2, WeightKg = 0.5 },
			new FunkyAnimal { Id = 3, Name = "Goat", Leg_Count = 4, WeightKg = 30 }
		};

	internal static List<Car> GetCars()
		=> new()
		{
			new Car { Id = 1, Name = "Ford Prefect", WheelCount = 4, WeightKg = 75 },
			new Car { Id = 2, Name = "Ford! Focus!", WheelCount = 4, WeightKg = 2000 }
		};

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

			Assert.Throws<InvalidOperationException>(() => loadMagicSpreadsheet.GetList<Car>("Animals"));

			// Try to load FunkyAnimals into a list of Animals (should succeed)
			loadMagicSpreadsheet.GetList<Animal>("Animals");
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
		var deviceSpecifications = sheet.GetExtendedList<object>("Successes");
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
		siteItem.Should().NotBeNull();
		siteItem!.FloorHeightFeet.Should().NotBe(0);
		sites.Should().NotBeEmpty();

		var devices = sheet.GetExtendedList<ImportedDevice>("Devices");
		devices.Should().NotBeEmpty();
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
		sites.Should().NotBeEmpty();

		var devices = sheet.GetExtendedList<ImportedDevice>("Devices");
		devices.Should().NotBeEmpty();
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

		Assert.ThrowsAny<Exception>(() => sheet.GetExtendedList<ImportedDevice>("XXX"));
	}

	[Fact]
	public void LoadSpreadsheet()
	{
		using var sheet = new MagicSpreadsheet(GetSheetFileInfo("Bulk Import Template"));
		sheet.Load();
		var deviceSpecifications = sheet.GetExtendedList<DeviceSpecification>();
		// do some sheet
		Assert.NotEmpty(deviceSpecifications);
		Assert.Single(deviceSpecifications);

		var device = deviceSpecifications[0];
		device.Item.Should().NotBeNull();
		device.Item!.HostName.Should().Be("localhost");
		device.Item.DeviceDisplayName.Should().Be("DeviceDisplayName");
		device.Item.DeviceDescription.Should().Be("The device description");
		device.Item.DeviceGroups.Should().Be("Group/SubGroup1;Group/SubGroup2");
		device.Item.PreferredCollector.Should().Be("CollectorDescription");
		device.Item.EnableAlerts.Should().Be(true);
		device.Item.EnableNetflow.Should().Be(false);
		device.Item.NetflowCollector.Should().Be(string.Empty);
		device.Item.Link.Should().Be("http://www.logicmonitor.com/");

		// make sure there are 2 custom properties and are the values we're expecting
		device.Properties.Count.Should().Be(2);
		device.Properties["Column A"].Should().Be("ValueA");
		device.Properties["column.b"].Should().Be("ValueB");
	}
}