namespace PanoramicData.SheetMagic.Test;

public class StringTests : Test
{
	[Theory]
	[InlineData("wmi.pass", "wmipass")]
	[InlineData("wmi pass", "wmipass")]
	[InlineData("wmi - (pass)", "wmipass")]
	[InlineData("1abc", "abc")]
	[InlineData("abc2", "abc2")]
	public void TweakStrings(string input, string expectedOutput)
		=> MagicSpreadsheet.TweakString(input).Should().Be(expectedOutput);
}