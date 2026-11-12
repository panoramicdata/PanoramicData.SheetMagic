namespace PanoramicData.SheetMagic.Test;

public partial class SaveSheetTests : Test
{
	private static Dictionary<string, object?> RoundTripProperties(Dictionary<string, object?> properties)
	{
		var a = new Extended<object>(new object(), properties);
		var fileInfo = GetXlsxTempFileInfo();

		try
		{
			using (var s1 = new MagicSpreadsheet(fileInfo))
			{
				s1.AddSheet([a]);
				s1.Save();
			}

			using var s2 = new MagicSpreadsheet(fileInfo);
			s2.Load();
			var b = s2.GetExtendedList<object>();
			_ = b.Should().NotBeNullOrEmpty();
			_ = b[0].Properties.Keys.Should().Contain("a");
			return b[0].Properties;
		}
		finally
		{
			fileInfo.Delete();
		}
	}

	private static object? RoundTripValue<T>(T value, bool includeRandomString = true)
	{
		var properties = new Dictionary<string, object?> { { "a", value } };
		if (includeRandomString)
		{
			properties.Add("b", "randomString");
		}

		return RoundTripProperties(properties)["a"];
	}

	private static void Check<T>(T theValue)
		=> _ = RoundTripValue(theValue).Should().Be(theValue);

	//When creating cells with DateTimeOffset we use the .UtcDateTime
	//which creates a DateTime object so we need to compare that
	private static void CheckDateTimeOffset(DateTimeOffset? inputDate, DateTime expectedDate)
		=> _ = RoundTripValue(inputDate, includeRandomString: false).Should().Be(expectedDate);

	private static void CheckNullValue<T>(T inputValue, object expectedValue)
		=> _ = RoundTripValue(inputValue).Should().Be(expectedValue);

	private static void CheckWithDoublePrecision<T>(T theValue) where T : struct
	{
		var retrievedValue = RoundTripValue(theValue);

		// Excel stores all numbers as doubles, so large integers lose precision
		_ = retrievedValue.Should().NotBeNull();
		_ = retrievedValue.Should().BeOfType<double>();

		var doubleValue = (double)retrievedValue!;
		var expectedDouble = Convert.ToDouble(theValue);
		_ = Math.Abs(doubleValue - expectedDouble).Should().BeLessThan(expectedDouble * 1e-10);
	}
}
