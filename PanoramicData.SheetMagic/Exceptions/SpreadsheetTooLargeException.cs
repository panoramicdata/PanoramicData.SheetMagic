namespace PanoramicData.SheetMagic.Exceptions;

/// <summary>
/// Exception thrown when a workbook is too large to be serialised.
/// </summary>
/// <remarks>
/// The underlying OpenXML writer serialises through a stream whose length is limited to
/// <see cref="int.MaxValue"/> bytes (approximately 2 GB).  Exceeding that limit surfaces from
/// the OpenXML library as a low-level <see cref="IOException"/> ("Stream was too long."), which
/// gives the caller no indication of the cause.  <see cref="MagicSpreadsheet.Save"/> translates
/// that case into this exception so it can be caught and diagnosed specifically.
/// </remarks>
public class SpreadsheetTooLargeException : SheetMagicException
{
	/// <summary>
	/// The maximum number of bytes that a serialised workbook may occupy.
	/// </summary>
	public const long MaximumSizeBytes = int.MaxValue;

	/// <summary>
	/// Initializes a new instance of the <see cref="SpreadsheetTooLargeException"/> class.
	/// </summary>
	public SpreadsheetTooLargeException()
		: base(BuildMessage())
	{
	}

	/// <summary>
	/// Initializes a new instance with the exception raised by the underlying OpenXML writer.
	/// </summary>
	/// <param name="innerException">The exception raised while serialising the workbook.</param>
	public SpreadsheetTooLargeException(Exception innerException)
		: base(BuildMessage(), innerException)
	{
	}

	/// <summary>
	/// Initializes a new instance with a message.
	/// </summary>
	/// <param name="message">The exception message.</param>
	public SpreadsheetTooLargeException(string message) : base(message)
	{
	}

	/// <summary>
	/// Initializes a new instance with a message and inner exception.
	/// </summary>
	/// <param name="message">The exception message.</param>
	/// <param name="innerException">The inner exception.</param>
	public SpreadsheetTooLargeException(string message, Exception innerException) : base(message, innerException)
	{
	}

	private static string BuildMessage()
		=> $"The workbook could not be saved because its serialised size exceeds the maximum supported size of {MaximumSizeBytes} bytes. Reduce the number of rows, columns or sheets written to the workbook.";
}
