﻿using Newtonsoft.Json.Linq;
using PanoramicData.SheetMagic.Test.Models;
using System;
using System.Collections.Generic;
using Xunit;

namespace PanoramicData.SheetMagic.Test;

public class AddSheetTests : Test
{
	[Theory]
	[InlineData("12345678901234567890123456789012")]
	[InlineData("abcdefghijklmnopqrstuvwxyz123456")]
	public void AddSheet_SheetNameTooLong_Fails(string badSheetName)
	{
		var fileInfo = GetXlsxTempFileInfo();
		var items = new List<SimpleAnimal> { new SimpleAnimal { Id = 1, Name = "Alligator" } };

		try
		{
			using var s = new MagicSpreadsheet(fileInfo);
			Assert.Throws<ArgumentException>(() => s.AddSheet(items, badSheetName));
		}
		finally
		{
			fileInfo.Delete();
		}
	}

	[Fact]
	public void AddSheet_JObjects_Succeeds()
	{
		var fileInfo = GetXlsxTempFileInfo();

		try
		{
			using var s = new MagicSpreadsheet(fileInfo);
			var jObjectList = new List<JObject>
				{
					JObject.FromObject(new SimpleAnimal { Id = 1, Name = "alligator" }),
					JObject.FromObject(new SimpleAnimal { Id = 2, Name = "bee" })
				};

			// Convert JObjects to Extended<object>
			var extendedList = new List<Extended<object>>();
			foreach (var jObject in jObjectList)
			{
				var extended = new Extended<object>(new(), jObject.ToObject<Dictionary<string, object?>>() ?? throw new ArgumentException("Could not convert JObject to dictionary"));
				extendedList.Add(extended);
			}

			s.AddSheet(
				extendedList
			);
		}
		finally
		{
			fileInfo.Delete();
		}
	}

	[Fact]
	public void AddSheet_SheetNameAlreadyExists_Fails()
	{
		var fileInfo = GetXlsxTempFileInfo();

		try
		{
			using var s = new MagicSpreadsheet(fileInfo);
			var items = new List<SimpleAnimal> { new SimpleAnimal { Id = 1, Name = "Alligator" } };
			s.AddSheet(items, "Sheet1");
			Assert.Throws<ArgumentException>(() => s.AddSheet(items, "Sheet1"));
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
					  new Extended<object>(new object(), new Dictionary<string, object?> {
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
					new FunkyAnimal{ Id = 0, Name = "Old Woman", WeightKg = 60, Leg_Count = 2},
					new FunkyAnimal{ Id = 1, Name = "Horse", WeightKg = 200, Leg_Count = 4},
					new FunkyAnimal{ Id = 2, Name = "Cow", WeightKg = 100, Leg_Count = 4},
					new FunkyAnimal{ Id = 3, Name = "Dog", WeightKg = 50, Leg_Count = 4},
					new FunkyAnimal{ Id = 4, Name = "Cat", WeightKg = 25, Leg_Count = 4},
					new FunkyAnimal{ Id = 5, Name = "Mouse", WeightKg = 0.1, Leg_Count = 4},
					new FunkyAnimal{ Id = 7, Name = "Spider", WeightKg = 0.01, Leg_Count = 8},
					new FunkyAnimal{ Id = 8, Name = "Fly", WeightKg = 0.001, Leg_Count = 6}
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
				new FunkyAnimal{ Id = 0, Name = "Old Woman", WeightKg = 60, Leg_Count = 2},
				new FunkyAnimal{ Id = 1, Name = "Horse", WeightKg = 200, Leg_Count = 4},
				new FunkyAnimal{ Id = 2, Name = "Cow", WeightKg = 100, Leg_Count = 4},
				new FunkyAnimal{ Id = 3, Name = "Dog", WeightKg = 50, Leg_Count = 4},
				new FunkyAnimal{ Id = 4, Name = "Cat", WeightKg = 25, Leg_Count = 4},
				new FunkyAnimal{ Id = 5, Name = "Mouse", WeightKg = 0.1, Leg_Count = 4},
				new FunkyAnimal{ Id = 7, Name = "Spider", WeightKg = 0.01, Leg_Count = 8},
				new FunkyAnimal{ Id = 8, Name = "Fly", WeightKg = 0.001, Leg_Count = 6}
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

			var funkyAnimals = new List<FunkyAnimal>
			{
				new FunkyAnimal{
					Id = 0,
					Name = "Old Woman",
					WeightKg = 60,
					Leg_Count = 2,
					Nicknames = new() {"Woo", "Yay"},
					Friends = new() {
						new() { Id = 10, Name="Houpla" },
						new() { Id = 11, Name = "Bedoink" }
					}
				},
				new FunkyAnimal{ Id = 1, Name = "Horse", WeightKg = 200, Leg_Count = 4, Nicknames = new() {"Bert", "Ernie" }},
				new FunkyAnimal{ Id = 2, Name = "Cow", WeightKg = 100, Leg_Count = 4},
				new FunkyAnimal{ Id = 3, Name = "Dog", WeightKg = 50, Leg_Count = 4},
				new FunkyAnimal{ Id = 4, Name = "Cat", WeightKg = 25, Leg_Count = 4},
				new FunkyAnimal{ Id = 5, Name = "Mouse", WeightKg = 0.1, Leg_Count = 4},
				new FunkyAnimal{ Id = 7, Name = "Spider", WeightKg = 0.01, Leg_Count = 8},
				new FunkyAnimal{ Id = 8, Name = "Fly", WeightKg = 0.001, Leg_Count = 6}
			};

			var sheetOptions = new AddSheetOptions
			{
				EnumerableCellOptions = new()
				{
					Expand = true,
					CellDelimiter = ", "
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
	public void AddSheet_EnumerablePropertyWithExtended_Succeeds()
	{
		var fileInfo = GetXlsxTempFileInfo();

		try
		{
			using var s = new MagicSpreadsheet(fileInfo);
			var emptyDictionary = new Dictionary<string, object?>();
			var funkyAnimals = new List<Extended<FunkyAnimal>>
			{
				new(new FunkyAnimal{
					Id = 0,
					Name = "Old Woman",
					WeightKg = 60,
					Leg_Count = 2,
					Nicknames = new() {"Woo", "Yay"},
					Friends = new() {
						new() { Id = 10, Name="Houpla" },
						new() { Id = 11, Name = "Bedoink" }
					},
				}, new Dictionary<string, object?> { { "Extended", "Extended" } }),
				new (
					new FunkyAnimal
					{
						Id = 1,
						Name = "Horse",
						WeightKg = 200,
						Leg_Count = 4,
						Nicknames = new() {"Bert", "Ernie" }
					},
					new Dictionary<string, object?>{
					}),
				new (new FunkyAnimal{ Id = 2, Name = "Cow", WeightKg = 100, Leg_Count = 4}, emptyDictionary),
				new (new FunkyAnimal{ Id = 3, Name = "Dog", WeightKg = 50, Leg_Count = 4}, emptyDictionary),
				new (new FunkyAnimal{ Id = 4, Name = "Cat", WeightKg = 25, Leg_Count = 4}, emptyDictionary),
				new (new FunkyAnimal{ Id = 5, Name = "Mouse", WeightKg = 0.1, Leg_Count = 4}, emptyDictionary),
				new (new FunkyAnimal{ Id = 7, Name = "Spider", WeightKg = 0.01, Leg_Count = 8}, emptyDictionary),
				new(new FunkyAnimal { Id = 8, Name = "Fly", WeightKg = 0.001, Leg_Count = 6 }, emptyDictionary)
			};

			var sheetOptions = new AddSheetOptions
			{
				EnumerableCellOptions = new()
				{
					Expand = true,
					CellDelimiter = ", "
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
}
