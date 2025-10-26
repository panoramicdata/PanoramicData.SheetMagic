using AwesomeAssertions;
using System.Collections.Generic;
using System.IO;
using Xunit;

namespace PanoramicData.SheetMagic.Test;

public class StreamTests
{
	[Fact]
	public void CreatingABinaryStream()
	{
		var a = new Extended<object>(
			new object(),
			new Dictionary<string, object?>
			{
				{ "a", "b" }
			}
		);
		using var stream = new MemoryStream();

		// Save
		using var s1 = new MagicSpreadsheet(stream);
		s1.AddSheet(new List<Extended<object>> { a }, "Sheet A");
		s1.AddSheet(new List<Extended<object>> { a }, "Sheet B");
		s1.Save();

		var bytes = stream.ToArray();

		bytes.Should().NotBeNullOrEmpty();
	}
}
