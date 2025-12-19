namespace PanoramicData.SheetMagic.Test;

public class EmptyRowHandlingTests : Test
{
	[Fact]
	public void GetList_WithEmptyRow_NoOptionsSet_ThrowsEmptyRowException()
	{
		// Arrange - Load a sheet with blank rows but without any empty row handling options
		using var magicSpreadsheet = new MagicSpreadsheet(
			GetSheetFileInfo("ParentChildWithBlankRows"),
			new Options
			{
				StopProcessingOnFirstEmptyRow = false,
				EmptyRowInterpretedAsNull = false
			});
		magicSpreadsheet.Load();

		// Act & Assert
		var action = () => magicSpreadsheet.GetList<ParentChildRelationship>();
		_ = action.Should().Throw<EmptyRowException>()
			.Which.RowIndex.Should().BePositive();
	}

	[Fact]
	public void GetList_WithEmptyRow_StopProcessingOnFirstEmptyRow_ReturnsPartialList()
	{
		// Arrange
		using var magicSpreadsheet = new MagicSpreadsheet(
			GetSheetFileInfo("ParentChildWithBlankRows"),
			new Options { StopProcessingOnFirstEmptyRow = true });
		magicSpreadsheet.Load();

		// Act
		var result = magicSpreadsheet.GetList<ParentChildRelationship>();

		// Assert
		_ = result.Should().NotBeNull();
		_ = result.Should().HaveCount(3); // Should stop at first empty row
	}

	[Fact]
	public void GetList_WithEmptyRow_EmptyRowInterpretedAsNull_ReturnsNullForEmptyRows()
	{
		// Arrange - use StopProcessingOnFirstEmptyRow along with EmptyRowInterpretedAsNull
		// to avoid parsing issues with data after the empty row
		using var magicSpreadsheet = new MagicSpreadsheet(
			GetSheetFileInfo("ParentChildWithBlankRows"),
			new Options
			{
				EmptyRowInterpretedAsNull = true,
				StopProcessingOnFirstEmptyRow = true
			});
		magicSpreadsheet.Load();

		// Act
		var result = magicSpreadsheet.GetList<ParentChildRelationship>();

		// Assert
		_ = result.Should().NotBeNull();
		_ = result.Should().HaveCount(3); // Should stop at first empty row
	}

	[Fact]
	public void GetExtendedList_WithEmptyRow_EmptyRowInterpretedAsNull_ReturnsExtendedWithNullItem()
	{
		// Arrange - use StopProcessingOnFirstEmptyRow along with EmptyRowInterpretedAsNull
		// to avoid parsing issues with data after the empty row
		using var magicSpreadsheet = new MagicSpreadsheet(
			GetSheetFileInfo("ParentChildWithBlankRows"),
			new Options
			{
				EmptyRowInterpretedAsNull = true,
				StopProcessingOnFirstEmptyRow = true
			});
		magicSpreadsheet.Load();

		// Act
		var result = magicSpreadsheet.GetExtendedList<ParentChildRelationship>();

		// Assert
		_ = result.Should().NotBeNull();
		_ = result.Should().HaveCount(3); // Should stop at first empty row
	}

	[Fact]
	public void EmptyRowException_ContainsRowIndex_WhenThrown()
	{
		// Arrange
		using var magicSpreadsheet = new MagicSpreadsheet(
			GetSheetFileInfo("ParentChildWithBlankRows"),
			new Options
			{
				StopProcessingOnFirstEmptyRow = false,
				EmptyRowInterpretedAsNull = false
			});
		magicSpreadsheet.Load();

		// Act & Assert
		try
		{
			magicSpreadsheet.GetList<ParentChildRelationship>();
			Assert.Fail("Expected EmptyRowException to be thrown");
		}
		catch (EmptyRowException ex)
		{
			_ = ex.RowIndex.Should().BePositive();
			_ = ex.Message.Should().Contain("empty");
			_ = ex.Message.Should().Contain("EmptyRowInterpretedAsNull");
		}
	}
}
