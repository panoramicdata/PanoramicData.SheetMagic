using PanoramicData.SheetMagic.Test.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace PanoramicData.SheetMagic.Test
{
    public class AddSheet : Test
    {
        [Theory]
        [InlineData("12345678901234567890123456789012")]
        [InlineData("abcdefghijklmnopqrstuvwxyz123456")]
        public void AddSheet_SheetNameTooLong_Fails(string badSheetName)
        {
            var fileInfo = GetXlsxTempFileInfo();

            try
            {
                using var s = new MagicSpreadsheet(fileInfo);
                Assert.Throws<ArgumentException>(() => s.AddSheet(new List<SimpleAnimal>(), badSheetName));
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
                s.AddSheet(new List<SimpleAnimal>(), "Sheet1");
                Assert.Throws<ArgumentException>(() => s.AddSheet(new List<SimpleAnimal>(), "Sheet1"));
            }
            finally
            {
                fileInfo.Delete();
            }
        }

        [Fact]
        public void AddSheet_SheetWithExtraProperties_Succeeds()
        {
            var fileInfo = GetXlsxTempFileInfo();

            try
            {
                using var s = new MagicSpreadsheet(fileInfo);
                var magicSheetListWithExtraProperties = new List<Extended<object>>();
                var propertyInfos = typeof(SimpleAnimal).GetProperties();

                var animals = new List<SimpleAnimal>
                {
                    new SimpleAnimal()
                    {
                        Id = 1,
                        Name = "Bob"
                    }
                };

                foreach (var animal in animals)
                {
                    var extendedObject = new Extended<object>()
                    {
                        Properties = new Dictionary<string, object>()
                    };

                    var propertyInfo = propertyInfos.Single(a => a.Name == nameof(SimpleAnimal.Id));

                    extendedObject.Properties.Add("My: column", propertyInfo.GetValue(animal));

                    // Store the extended object
                    magicSheetListWithExtraProperties.Add(extendedObject);
                }

                s.AddSheet(magicSheetListWithExtraProperties, "Sheet1");
                s.Save();
            }
            finally
            {
                fileInfo.Delete();
            }
        }
    }
}
