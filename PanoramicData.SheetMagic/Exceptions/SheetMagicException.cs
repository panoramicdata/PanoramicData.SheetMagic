using System;
using System.Runtime.Serialization;

namespace PanoramicData.SheetMagic.Exceptions
{
	public abstract class SheetMagicException : Exception
	{
		internal SheetMagicException()
		{
		}

		internal SheetMagicException(string message) : base(message)
		{
		}

		internal SheetMagicException(string message, Exception innerException) : base(message, innerException)
		{
		}

		internal SheetMagicException(SerializationInfo info, StreamingContext context) : base(info, context)
		{
		}
	}
}