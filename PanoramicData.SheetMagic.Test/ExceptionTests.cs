namespace PanoramicData.SheetMagic.Test;

public class ExceptionTests
{
	#region EmptyRowException Tests

	[Fact]
	public void EmptyRowException_DefaultConstructor_CreatesInstance()
	{
		// Act
		var exception = new EmptyRowException();

		// Assert
		_ = exception.Should().NotBeNull();
		_ = exception.RowIndex.Should().Be(0);
		_ = exception.Message.Should().NotBeNullOrEmpty();
	}

	[Fact]
	public void EmptyRowException_WithRowIndex_SetsRowIndexAndMessage()
	{
		// Arrange
		const int rowIndex = 5;

		// Act
		var exception = new EmptyRowException(rowIndex);

		// Assert
		_ = exception.RowIndex.Should().Be(rowIndex);
		_ = exception.Message.Should().Contain("5");
		_ = exception.Message.Should().Contain("empty");
	}

	[Fact]
	public void EmptyRowException_WithMessage_SetsMessage()
	{
		// Arrange
		const string message = "Custom error message";

		// Act
		var exception = new EmptyRowException(message);

		// Assert
		_ = exception.Message.Should().Be(message);
	}

	[Fact]
	public void EmptyRowException_WithMessageAndInnerException_SetsProperties()
	{
		// Arrange
		const string message = "Custom error message";
		var innerException = new InvalidOperationException("Inner error");

		// Act
		var exception = new EmptyRowException(message, innerException);

		// Assert
		_ = exception.Message.Should().Be(message);
		_ = exception.InnerException.Should().Be(innerException);
	}

	[Fact]
	public void EmptyRowException_InheritsFromSheetMagicException()
	{
		// Act
		var exception = new EmptyRowException(1);

		// Assert
		_ = exception.Should().BeAssignableTo<SheetMagicException>();
	}

	#endregion

	#region PropertyNotFoundException Tests

	[Fact]
	public void PropertyNotFoundException_DefaultConstructor_CreatesInstance()
	{
		// Act
		var exception = new PropertyNotFoundException();

		// Assert
		_ = exception.Should().NotBeNull();
		_ = exception.PropertyName.Should().BeEmpty();
	}

	[Fact]
	public void PropertyNotFoundException_WithPropertyName_SetsPropertyNameAndMessage()
	{
		// Arrange
		const string propertyName = "TestProperty";

		// Act
		var exception = new PropertyNotFoundException(propertyName);

		// Assert
		_ = exception.PropertyName.Should().Be(propertyName);
		_ = exception.Message.Should().Contain(propertyName);
		_ = exception.Message.Should().Contain("not found");
	}

	[Fact]
	public void PropertyNotFoundException_WithPropertyNameAndMessage_SetsProperties()
	{
		// Arrange
		const string propertyName = "TestProperty";
		const string message = "Custom message";

		// Act
		var exception = new PropertyNotFoundException(propertyName, message);

		// Assert
		_ = exception.PropertyName.Should().Be(propertyName);
		_ = exception.Message.Should().Be(message);
	}

	[Fact]
	public void PropertyNotFoundException_WithPropertyNameAndInnerException_SetsProperties()
	{
		// Arrange
		const string propertyName = "TestProperty";
		var innerException = new InvalidOperationException("Inner error");

		// Act
		var exception = new PropertyNotFoundException(propertyName, innerException);

		// Assert
		_ = exception.PropertyName.Should().Be(propertyName);
		_ = exception.InnerException.Should().Be(innerException);
		_ = exception.Message.Should().Contain(propertyName);
	}

	[Fact]
	public void PropertyNotFoundException_WithAllParameters_SetsAllProperties()
	{
		// Arrange
		const string propertyName = "TestProperty";
		const string message = "Custom message";
		var innerException = new InvalidOperationException("Inner error");

		// Act
		var exception = new PropertyNotFoundException(propertyName, message, innerException);

		// Assert
		_ = exception.PropertyName.Should().Be(propertyName);
		_ = exception.Message.Should().Be(message);
		_ = exception.InnerException.Should().Be(innerException);
	}

	[Fact]
	public void PropertyNotFoundException_InheritsFromException()
	{
		// Act
		var exception = new PropertyNotFoundException("Test");

		// Assert
		_ = exception.Should().BeAssignableTo<Exception>();
	}

	#endregion

	#region ValidationException Tests

	[Fact]
	public void ValidationException_DefaultConstructor_CreatesInstance()
	{
		// Act
		var exception = new ValidationException();

		// Assert
		_ = exception.Should().NotBeNull();
	}

	[Fact]
	public void ValidationException_WithMessage_SetsMessage()
	{
		// Arrange
		const string message = "Validation failed";

		// Act
		var exception = new ValidationException(message);

		// Assert
		_ = exception.Message.Should().Be(message);
	}

	[Fact]
	public void ValidationException_WithMessageAndInnerException_SetsProperties()
	{
		// Arrange
		const string message = "Validation failed";
		var innerException = new InvalidOperationException("Inner error");

		// Act
		var exception = new ValidationException(message, innerException);

		// Assert
		_ = exception.Message.Should().Be(message);
		_ = exception.InnerException.Should().Be(innerException);
	}

	[Fact]
	public void ValidationException_InheritsFromSheetMagicException()
	{
		// Act
		var exception = new ValidationException("Test");

		// Assert
		_ = exception.Should().BeAssignableTo<SheetMagicException>();
	}

	#endregion
}
