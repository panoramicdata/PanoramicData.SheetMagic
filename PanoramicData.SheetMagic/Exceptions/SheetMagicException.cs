namespace PanoramicData.SheetMagic.Exceptions;

/// <summary>
/// Base exception class for all SheetMagic-related exceptions.
/// </summary>
public abstract class SheetMagicException : Exception
{
	/// <summary>
	/// Initializes a new instance of the <see cref="SheetMagicException"/> class.
	/// </summary>
	protected SheetMagicException()
	{
	}

	/// <summary>
	/// Initializes a new instance of the <see cref="SheetMagicException"/> class with a message.
	/// </summary>
	/// <param name="message">The exception message.</param>
	protected SheetMagicException(string message) : base(message)
	{
	}

	/// <summary>
	/// Initializes a new instance of the <see cref="SheetMagicException"/> class with a message and inner exception.
	/// </summary>
	/// <param name="message">The exception message.</param>
	/// <param name="innerException">The inner exception.</param>
	protected SheetMagicException(string message, Exception innerException) : base(message, innerException)
	{
	}
}