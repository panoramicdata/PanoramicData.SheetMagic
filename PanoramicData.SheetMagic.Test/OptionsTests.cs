using AwesomeAssertions;
using PanoramicData.SheetMagic.Exceptions;
using Xunit;

namespace PanoramicData.SheetMagic.Test;

public class OptionsTests
{
	#region Options Tests

	[Fact]
	public void Options_DefaultValues_AreCorrect()
	{
		// Act
		var options = new Options();

		// Assert
		_ = options.StopProcessingOnFirstEmptyRow.Should().BeFalse();
		_ = options.EmptyRowInterpretedAsNull.Should().BeFalse();
		_ = options.IgnoreUnmappedProperties.Should().BeFalse();
		_ = options.LoadNullExtendedProperties.Should().BeFalse();
		_ = options.IsLoadedFileEditable.Should().BeFalse();
		_ = options.ListSeparator.Should().Be(", ");
		_ = options.DefaultAddSheetOptions.Should().NotBeNull();
		_ = options.TableStyles.Should().NotBeNull();
		_ = options.TableStyles.Should().BeEmpty();
		_ = options.EnumerableCellOptions.Should().NotBeNull();
	}

	[Fact]
	public void Options_SetProperties_RetainsValues()
	{
		// Arrange & Act
		var options = new Options
		{
			StopProcessingOnFirstEmptyRow = true,
			EmptyRowInterpretedAsNull = true,
			IgnoreUnmappedProperties = true,
			LoadNullExtendedProperties = true,
			IsLoadedFileEditable = true,
			ListSeparator = "; "
		};

		// Assert
		_ = options.StopProcessingOnFirstEmptyRow.Should().BeTrue();
		_ = options.EmptyRowInterpretedAsNull.Should().BeTrue();
		_ = options.IgnoreUnmappedProperties.Should().BeTrue();
		_ = options.LoadNullExtendedProperties.Should().BeTrue();
		_ = options.IsLoadedFileEditable.Should().BeTrue();
		_ = options.ListSeparator.Should().Be("; ");
	}

	[Fact]
	public void Options_Validate_WithEmptyTableStyles_Succeeds()
	{
		// Arrange
		var options = new Options();

		// Act & Assert - should not throw
		options.Validate();
	}

	[Fact]
	public void Options_Validate_WithValidTableStyles_Succeeds()
	{
		// Arrange
		var options = new Options
		{
			TableStyles =
			[
				new CustomTableStyle
				{
					Name = "TestStyle",
					HeaderRowStyle = new TableRowStyle { BackgroundColor = System.Drawing.Color.Blue }
				}
			]
		};

		// Act & Assert - should not throw
		options.Validate();
	}

	[Fact]
	public void Options_Validate_WithInvalidTableStyle_ThrowsValidationException()
	{
		// Arrange
		var options = new Options
		{
			TableStyles =
			[
				new CustomTableStyle { Name = "" }
			]
		};

		// Act & Assert
		var action = () => options.Validate();
		_ = action.Should().Throw<ValidationException>();
	}

	#endregion

	#region CustomTableStyle Tests

	[Fact]
	public void CustomTableStyle_DefaultValues_AreCorrect()
	{
		// Act
		var style = new CustomTableStyle();

		// Assert
		_ = style.Name.Should().Be("Custom Table Style");
		_ = style.HeaderRowStyle.Should().BeNull();
		_ = style.OddRowStyle.Should().BeNull();
		_ = style.EvenRowStyle.Should().BeNull();
		_ = style.WholeTableStyle.Should().BeNull();
	}

	[Fact]
	public void CustomTableStyle_SetProperties_RetainsValues()
	{
		// Arrange
		var headerStyle = new TableRowStyle { BackgroundColor = System.Drawing.Color.Blue };
		var oddStyle = new TableRowStyle { BackgroundColor = System.Drawing.Color.LightGray };
		var evenStyle = new TableRowStyle { BackgroundColor = System.Drawing.Color.White };
		var wholeTableStyle = new TableRowStyle { FontColor = System.Drawing.Color.Black };

		// Act
		var style = new CustomTableStyle
		{
			Name = "MyStyle",
			HeaderRowStyle = headerStyle,
			OddRowStyle = oddStyle,
			EvenRowStyle = evenStyle,
			WholeTableStyle = wholeTableStyle
		};

		// Assert
		_ = style.Name.Should().Be("MyStyle");
		_ = style.HeaderRowStyle.Should().Be(headerStyle);
		_ = style.OddRowStyle.Should().Be(oddStyle);
		_ = style.EvenRowStyle.Should().Be(evenStyle);
		_ = style.WholeTableStyle.Should().Be(wholeTableStyle);
	}

	[Fact]
	public void CustomTableStyle_Validate_WithEmptyName_ThrowsValidationException()
	{
		// Arrange
		var style = new CustomTableStyle
		{
			Name = "",
			HeaderRowStyle = new TableRowStyle()
		};

		// Act & Assert
		var action = () => style.Validate();
		_ = action.Should().Throw<ValidationException>()
			.WithMessage("*no name*");
	}

	[Fact]
	public void CustomTableStyle_Validate_WithWhitespaceName_ThrowsValidationException()
	{
		// Arrange
		var style = new CustomTableStyle
		{
			Name = "   ",
			HeaderRowStyle = new TableRowStyle()
		};

		// Act & Assert
		var action = () => style.Validate();
		_ = action.Should().Throw<ValidationException>()
			.WithMessage("*no name*");
	}

	[Fact]
	public void CustomTableStyle_Validate_WithNoStyles_ThrowsValidationException()
	{
		// Arrange
		var style = new CustomTableStyle
		{
			Name = "ValidName"
		};

		// Act & Assert
		var action = () => style.Validate();
		_ = action.Should().Throw<ValidationException>()
			.WithMessage("*No style set*");
	}

	[Fact]
	public void CustomTableStyle_Validate_WithHeaderRowStyleOnly_Succeeds()
	{
		// Arrange
		var style = new CustomTableStyle
		{
			Name = "ValidName",
			HeaderRowStyle = new TableRowStyle { BackgroundColor = System.Drawing.Color.Blue }
		};

		// Act & Assert - should not throw
		style.Validate();
	}

	[Fact]
	public void CustomTableStyle_Validate_WithOddRowStyleOnly_Succeeds()
	{
		// Arrange
		var style = new CustomTableStyle
		{
			Name = "ValidName",
			OddRowStyle = new TableRowStyle { BackgroundColor = System.Drawing.Color.LightGray }
		};

		// Act & Assert - should not throw
		style.Validate();
	}

	[Fact]
	public void CustomTableStyle_Validate_WithEvenRowStyleOnly_Succeeds()
	{
		// Arrange
		var style = new CustomTableStyle
		{
			Name = "ValidName",
			EvenRowStyle = new TableRowStyle { BackgroundColor = System.Drawing.Color.White }
		};

		// Act & Assert - should not throw
		style.Validate();
	}

	[Fact]
	public void CustomTableStyle_Validate_WithWholeTableStyleOnly_Succeeds()
	{
		// Arrange
		var style = new CustomTableStyle
		{
			Name = "ValidName",
			WholeTableStyle = new TableRowStyle { FontColor = System.Drawing.Color.Black }
		};

		// Act & Assert - should not throw
		style.Validate();
	}

	#endregion

	#region TableRowStyle Tests

	[Fact]
	public void TableRowStyle_DefaultValues_AreNull()
	{
		// Act
		var style = new TableRowStyle();

		// Assert
		_ = style.BackgroundColor.Should().BeNull();
		_ = style.FontColor.Should().BeNull();
		_ = style.InnerBorderColor.Should().BeNull();
		_ = style.OuterBorderColor.Should().BeNull();
		_ = style.FontWeight.Should().Be(FontWeight.Regular);
	}

	[Fact]
	public void TableRowStyle_SetProperties_RetainsValues()
	{
		// Act
		var style = new TableRowStyle
		{
			BackgroundColor = System.Drawing.Color.Blue,
			FontColor = System.Drawing.Color.White,
			InnerBorderColor = System.Drawing.Color.Gray,
			OuterBorderColor = System.Drawing.Color.Black,
			FontWeight = FontWeight.Bold
		};

		// Assert
		_ = style.BackgroundColor.Should().Be(System.Drawing.Color.Blue);
		_ = style.FontColor.Should().Be(System.Drawing.Color.White);
		_ = style.InnerBorderColor.Should().Be(System.Drawing.Color.Gray);
		_ = style.OuterBorderColor.Should().Be(System.Drawing.Color.Black);
		_ = style.FontWeight.Should().Be(FontWeight.Bold);
	}

	#endregion
}
