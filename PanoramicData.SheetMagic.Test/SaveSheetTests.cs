namespace PanoramicData.SheetMagic.Test;

public partial class SaveSheetTests : Test
{
	[Fact]
	public void SaveSheet_WithNoData_Succeeds()
	{
		var fileInfo = GetXlsxTempFileInfo();

		try
		{
			using var s = new MagicSpreadsheet(fileInfo);
			s.Save();
		}
		finally
		{
			fileInfo.Delete();
		}
	}

	[Fact]
	public void SavingWithExtendedObject_Succeeds()
	{
		var theString = "b";

		Check(theString);
	}

	[Theory]
	[InlineData(int.MinValue)]
	[InlineData(-1)]
	[InlineData(0)]
	[InlineData(1)]
	[InlineData(int.MaxValue)]
	public void SavingWithExtendedObjectContainingInt32_Succeeds(long value)
		=> Check(value);

	[Theory]
	[InlineData((uint)0)]
	[InlineData((uint)1)]
	[InlineData(uint.MaxValue)]
	public void SavingWithExtendedObjectContainingUInt32_Succeeds(uint value)
		=> Check(value);

	[Theory]
	[InlineData(long.MinValue)]
	[InlineData((long)-1)]
	[InlineData((long)0)]
	[InlineData((long)1)]
	public void SavingWithExtendedObjectContainingInt64_Succeeds(long value)
		=> Check(value);

	[Fact]
	public void SavingWithExtendedObjectContainingInt64_MaxValue_LosesPrecision() =>
		// Excel stores all numbers as doubles, which can't represent Int64.MaxValue with full precision
		// The value gets converted to scientific notation: 9.223372036854776E+18
		CheckWithDoublePrecision(long.MaxValue);

	[Theory]
	[InlineData((ulong)0)]
	[InlineData((ulong)1)]
	public void SavingWithExtendedObjectContainingUInt64_Succeeds(ulong value)
		=> Check(value);

	[Fact]
	public void SavingWithExtendedObjectContainingUInt64_MaxValue_LosesPrecision() =>
		// Excel stores all numbers as doubles, which can't represent UInt64.MaxValue with full precision
		// The value gets converted to scientific notation: 1.8446744073709552E+19
		CheckWithDoublePrecision(ulong.MaxValue);

	[Theory]
	[InlineData(short.MinValue)]
	[InlineData((short)-1)]
	[InlineData((short)0)]
	[InlineData((short)1)]
	[InlineData(short.MaxValue)]
	public void SavingWithExtendedObjectContainingInt16_Succeeds(short value)
		=> Check(value);

	[Theory]
	[InlineData((ushort)0)]
	[InlineData((ushort)1)]
	[InlineData(ushort.MaxValue)]
	public void SavingWithExtendedObjectContainingUInt16_Succeeds(ushort value)
		=> Check(value);

	[Theory]
	[InlineData(true)]
	[InlineData(false)]
	public void SavingWithExtendedObjectContainingBoolean_Succeeds(bool inputBool)
		=> Check(inputBool);

	[Theory]
	[InlineData(true)]
	[InlineData(false)]
	[InlineData(null)]
	public void SavingWithExtendedObjectContainingNullableBoolean_Succeeds(bool? inputBool)
	{
		if (inputBool == null)
		{
			// Excel stores null booleans as empty strings
			CheckNullValue<bool?>(inputBool, string.Empty);
		}
		else
		{
			Check(inputBool);
		}
	}

	[Theory]
	[InlineData(double.MinValue)]
	[InlineData(double.MaxValue)]
	[InlineData(double.NegativeInfinity)]
	[InlineData(double.PositiveInfinity)]
	[InlineData(12.3)]
	[InlineData(-1.0)]
	public void SavingWithExtendedObjectContainingDouble_Succeeds(double value)
		=> Check(value);

	[Fact]
	public void SavingWithExtendedObjectContainingDouble_NaN_Succeeds() =>
		// Excel stores NaN as empty string
		CheckNullValue(double.NaN, string.Empty);

	[Fact]
	public void SavingWithExtendedObjectContainingDateTime_Succeeds()
		=> Check(new DateTime(2000, 1, 2, 3, 4, 5));

	[Fact]
	public void SavingWithExtendedObjectContainingNullableDateTime_Succeeds()
		=> Check((DateTime?)new DateTime(2000, 1, 2, 3, 4, 5));

	[Fact]
	public void SavingWithExtendedObjectContainingDateTimeOffset_Succeeds()
		=> CheckDateTimeOffset(new DateTimeOffset(2000, 1, 2, 3, 4, 5, TimeSpan.Zero), new DateTimeOffset(2000, 1, 2, 3, 4, 5, TimeSpan.Zero).UtcDateTime);

	[Fact]
	public void SavingWithExtendedObjectContainingNullableDateTimeOffset_Succeeds()
		=> CheckDateTimeOffset((DateTimeOffset?)new DateTimeOffset(2000, 1, 2, 3, 4, 5, TimeSpan.Zero), new DateTimeOffset(2000, 1, 2, 3, 4, 5, TimeSpan.Zero).UtcDateTime);
}
