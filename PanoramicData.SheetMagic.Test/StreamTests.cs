using System.IO;

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
		s1.AddSheet([a], "Sheet A");
		s1.AddSheet([a], "Sheet B");
		s1.Save();

		var bytes = stream.ToArray();

		bytes.Should().NotBeNullOrEmpty();
	}
}
