using System.IO;
using System.Reflection;

namespace PanoramicData.SheetMagic.Test;

/// <summary>
/// Regression coverage for MS-24969: a <see cref="MagicSpreadsheet"/> whose save fails must not
/// be able to kill the host process from the finalizer thread.
/// </summary>
public class LifecycleTests : Test
{
	[Fact]
	public void MagicSpreadsheet_HasNoFinalizer()
	{
		// A finalizer on this type is the defect itself: it disposed the SpreadsheetDocument,
		// which saves, which can throw - and a throwing finalizer terminates the process.
		var finalizer = typeof(MagicSpreadsheet).GetMethod(
			"Finalize",
			BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.DeclaredOnly);

		_ = finalizer.Should().BeNull(
			"MagicSpreadsheet owns no unmanaged resources, and a finalizer that disposes the document can throw on the finalizer thread");
	}

	[Fact]
	public void Dispose_CalledTwice_DoesNotThrow()
	{
		var fileInfo = GetXlsxTempFileInfo();

		try
		{
			var spreadsheet = new MagicSpreadsheet(fileInfo);
			spreadsheet.AddSheet(new[] { new { Value = "a" } }.ToList());
			spreadsheet.Save();

			var dispose = () =>
			{
				spreadsheet.Dispose();
				spreadsheet.Dispose();
			};

			_ = dispose.Should().NotThrow();
		}
		finally
		{
			fileInfo.Refresh();
			if (fileInfo.Exists)
			{
				fileInfo.Delete();
			}
		}
	}

	[Fact]
	public void Dispose_AfterFailedSave_DoesNotRetryTheSave()
	{
		// Reproduces the production failure shape: the save fails, then the using block's
		// Dispose() runs.  Dispose() must not attempt the save again - previously it did, which
		// both replaced the original exception and left the instance finalizable.
		var fileInfo = new FileInfo(Path.Combine(
			Path.GetTempPath(),
			Guid.NewGuid().ToString(),
			"no-such-directory",
			"output.xlsx"));

		var spreadsheet = new MagicSpreadsheet(fileInfo);

		// Creating the document under a non-existent directory fails.
		_ = ((Action)(() => spreadsheet.AddSheet(new[] { new { Value = "a" } }.ToList())))
			.Should()
			.Throw<Exception>();

		_ = ((Action)spreadsheet.Dispose).Should().NotThrow();
	}

	[Fact]
	public void Save_CalledTwice_Throws()
	{
		var fileInfo = GetXlsxTempFileInfo();

		try
		{
			using var spreadsheet = new MagicSpreadsheet(fileInfo);
			spreadsheet.AddSheet(new[] { new { Value = "a" } }.ToList());
			spreadsheet.Save();

			_ = ((Action)spreadsheet.Save)
				.Should()
				.Throw<InvalidOperationException>()
				.WithMessage("*already been saved*");
		}
		finally
		{
			fileInfo.Refresh();
			if (fileInfo.Exists)
			{
				fileInfo.Delete();
			}
		}
	}

	[Fact]
	public void SpreadsheetTooLargeException_CarriesActionableDetail()
	{
		var inner = new IOException("Stream was too long.");

		var exception = new SpreadsheetTooLargeException(inner);

		_ = exception.Should().BeAssignableTo<SheetMagicException>();
		_ = exception.InnerException.Should().BeSameAs(inner);
		_ = exception.Message.Should().Contain(SpreadsheetTooLargeException.MaximumSizeBytes.ToString());
	}
}
