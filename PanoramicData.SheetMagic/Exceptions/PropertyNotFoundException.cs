namespace PanoramicData.SheetMagic.Exceptions;

/// <summary>
/// Exception thrown when a property cannot be found on a type.
/// </summary>
public class PropertyNotFoundException : Exception
{
	/// <summary>
	/// Initializes a new instance of the <see cref="PropertyNotFoundException"/> class.
	/// </summary>
	public PropertyNotFoundException()
	{
	}

	/// <summary>
	/// Initializes a new instance with the property name.
	/// </summary>
	/// <param name="propertyName">The name of the property that was not found.</param>
	public PropertyNotFoundException(string propertyName) : base($"Property '{propertyName}' not found")
	{
		PropertyName = propertyName;
	}

	/// <summary>
	/// Initializes a new instance with the property name and a custom message.
	/// </summary>
	/// <param name="propertyName">The name of the property that was not found.</param>
	/// <param name="message">The exception message.</param>
	public PropertyNotFoundException(string propertyName, string message) : base(message)
	{
		PropertyName = propertyName;
	}

	/// <summary>
	/// Initializes a new instance with the property name and inner exception.
	/// </summary>
	/// <param name="propertyName">The name of the property that was not found.</param>
	/// <param name="innerException">The inner exception.</param>
	public PropertyNotFoundException(string propertyName, Exception innerException) : base($"Property '{propertyName}' not found", innerException)
	{
		PropertyName = propertyName;
	}

	/// <summary>
	/// Initializes a new instance with the property name, message, and inner exception.
	/// </summary>
	/// <param name="propertyName">The name of the property that was not found.</param>
	/// <param name="message">The exception message.</param>
	/// <param name="innerException">The inner exception.</param>
	public PropertyNotFoundException(string propertyName, string message, Exception innerException) : base(message, innerException)
	{
		PropertyName = propertyName;
	}

	/// <summary>
	/// Gets the name of the property that was not found.
	/// </summary>
	public string PropertyName { get; private set; } = string.Empty;
}
