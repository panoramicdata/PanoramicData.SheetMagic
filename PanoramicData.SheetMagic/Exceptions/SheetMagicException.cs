namespace PanoramicData.SheetMagic.Exceptions;

public abstract class SheetMagicException : Exception
{
	protected SheetMagicException()
	{
	}

	protected SheetMagicException(string message) : base(message)
	{
	}

	protected SheetMagicException(string message, Exception innerException) : base(message, innerException)
	{
	}
}