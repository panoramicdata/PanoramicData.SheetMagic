namespace PanoramicData.SheetMagic.Interfaces;

/// <summary>
/// Interface for objects that can validate themselves.
/// </summary>
public interface IValidate
{
	/// <summary>
	/// Validates the object and throws an exception if invalid.
	/// </summary>
	/// <exception cref="Exceptions.ValidationException">Thrown when validation fails.</exception>
	void Validate();
}
