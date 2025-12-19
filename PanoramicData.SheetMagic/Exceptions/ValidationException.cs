namespace PanoramicData.SheetMagic.Exceptions;

/// <summary>
/// Exception thrown when validation of configuration or data fails.
/// </summary>
public class ValidationException : SheetMagicException
{
	/// <summary>
	/// Initializes a new instance of the <see cref="ValidationException"/> class.
	/// </summary>
	public ValidationException()
	{
	}

	/// <summary>
	/// Initializes a new instance with a message.
	/// </summary>
	/// <param name="message">The validation error message.</param>
	public ValidationException(string message) : base(message)
	{
	}

	/// <summary>
	/// Initializes a new instance with a message and inner exception.
	/// </summary>
	/// <param name="message">The validation error message.</param>
	/// <param name="innerException">The inner exception.</param>
	public ValidationException(string message, Exception innerException) : base(message, innerException)
	{
	}
}