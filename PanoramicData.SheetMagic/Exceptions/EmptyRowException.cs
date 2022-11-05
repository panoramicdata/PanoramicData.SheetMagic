using System;

namespace PanoramicData.SheetMagic.Exceptions;

public class EmptyRowException : SheetMagicException
{
	public int RowIndex { get; set; }

	public EmptyRowException(int rowIndex)
	  : base($"Row with index {rowIndex} is empty.  If this is permissible, set EmptyRowInterpretedAsNull or LoadNullExtendedProperties in the options.")
	{
		RowIndex = rowIndex;
	}

	public EmptyRowException()
	{
	}

	public EmptyRowException(string message) : base(message)
	{
	}

	public EmptyRowException(string message, Exception innerException) : base(message, innerException)
	{
	}
}
