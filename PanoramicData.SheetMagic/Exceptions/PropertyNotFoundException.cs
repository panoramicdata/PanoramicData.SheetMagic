namespace PanoramicData.SheetMagic.Exceptions;

public class PropertyNotFoundException : Exception
{
	public PropertyNotFoundException()
	{
	}

	public PropertyNotFoundException(string propertyName) : base($"Property '{propertyName}' not found")
	{
		PropertyName = propertyName;
	}

	public PropertyNotFoundException(string propertyName, string message) : base(message)
	{
		PropertyName = propertyName;
	}

	public PropertyNotFoundException(string propertyName, Exception innerException) : base($"Property '{propertyName}' not found", innerException)
	{
		PropertyName = propertyName;
	}

	public PropertyNotFoundException(string propertyName, string message, Exception innerException) : base(message, innerException)
	{
		PropertyName = propertyName;
	}

	public string PropertyName { get; private set; } = string.Empty;
}
