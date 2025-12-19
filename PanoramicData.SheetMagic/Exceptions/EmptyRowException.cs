namespace PanoramicData.SheetMagic.Exceptions;

/// <summary>
/// Exception thrown when an empty row is encountered during spreadsheet reading.
/// </summary>
public class EmptyRowException : SheetMagicException
{
	/// <summary>
	/// Gets or sets the index of the empty row.
	/// </summary>
	public int RowIndex { get; set; }

	/// <summary>
	/// Initializes a new instance with the row index.
	/// </summary>
	/// <param name="rowIndex">The index of the empty row.</param>
	public EmptyRowException(int rowIndex)
	  : base($"Row with index {rowIndex} is empty.  If this is permissible, set EmptyRowInterpretedAsNull or LoadNullExtendedProperties in the options.")
	{
		RowIndex = rowIndex;
	}

	/// <summary>
	/// Initializes a new instance of the <see cref="EmptyRowException"/> class.
	/// </summary>
	public EmptyRowException()
	{
	}

	/// <summary>
	/// Initializes a new instance with a message.
	/// </summary>
	/// <param name="message">The exception message.</param>
	public EmptyRowException(string message) : base(message)
	{
	}

	/// <summary>
	/// Initializes a new instance with a message and inner exception.
	/// </summary>
	/// <param name="message">The exception message.</param>
	/// <param name="innerException">The inner exception.</param>
	public EmptyRowException(string message, Exception innerException) : base(message, innerException)
	{
	}
}
