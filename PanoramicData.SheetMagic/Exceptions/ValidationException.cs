namespace PanoramicData.SheetMagic.Exceptions;

public class ValidationException : SheetMagicException
{
	public ValidationException()
	{
	}

	public ValidationException(string message) : base(message)
	{
	}

	public ValidationException(string message, Exception innerException) : base(message, innerException)
	{
	}
}