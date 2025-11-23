using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Sheet = DocumentFormat.OpenXml.Spreadsheet.Sheet;

namespace PanoramicData.SheetMagic;

/// <summary>
/// Core functionality for MagicSpreadsheet - constructors, fields, properties, and lifecycle methods
/// </summary>
public partial class MagicSpreadsheet : IDisposable
{
	private const string Letters = "abcdefghijklmnopqrstuvwxyz";
	private const string Numbers = "0123456789";
	private static readonly Regex CellReferenceRegex = GetCellReferenceRegex();

	private readonly FileInfo? _fileInfo;
	private readonly Stream? _stream;
	private readonly Options _options;
	private readonly HashSet<string> _uniqueTableDisplayNames = [];

	private SpreadsheetDocument? _document;

	public MagicSpreadsheet(FileInfo fileInfo, Options options)
	{
		_fileInfo = fileInfo;
		_options = options;
	}

	public MagicSpreadsheet(FileInfo fileInfo)
		: this(fileInfo, new Options())
	{
	}

	public MagicSpreadsheet(Stream stream, Options options)
	{
		_stream = stream;
		_options = options;
	}

	public MagicSpreadsheet(Stream stream)
		: this(stream, new Options())
	{
	}

	public List<string> SheetNames
		=> [.. ((((_document ?? throw new InvalidOperationException("No document loaded."))
			.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart not created"))
			.Workbook ?? throw new InvalidOperationException("Workbook not created"))
			.Sheets ?? throw new InvalidOperationException("Sheets not created"))
			.ChildElements
			.Cast<Sheet>()
			.Select(static s => s.Name?.Value ?? string.Empty)];

	public void Load() => _document = _fileInfo is not null
		? SpreadsheetDocument.Open(_fileInfo.FullName, false)
		: SpreadsheetDocument.Open(_stream!, false);

	public void Save()
	{
		// Ensure that at least one sheet has been added
		if (_document?.WorkbookPart?.Workbook?.Sheets == null || _document.WorkbookPart.Workbook.Sheets.Count() == 0)
		{
			// This has to contain some data to prevent file corruption.
			AddSheet(new[] { new { Error = "No data was output." } }.ToList(), "Sheet1");
		}

		if (_document?.WorkbookPart?.Workbook is null)
		{
			throw new InvalidOperationException("Document incorrectly created.");
		}

		_document.WorkbookPart.Workbook.Save();
		_document.Dispose();

		// Do we have a stream?
		if (_stream is not null)
		{
			// YES - Ensure it's flushed and seek back to the beginning for consumption
			_stream.Flush();
			_ = _stream.Seek(0, SeekOrigin.Begin);
		}
	}

	private void ReleaseUnmanagedResources() => _document?.Dispose();

	public void Dispose()
	{
		ReleaseUnmanagedResources();
		GC.SuppressFinalize(this);
	}

	~MagicSpreadsheet()
	{
		ReleaseUnmanagedResources();
	}

	[GeneratedRegex(@"(?<col>([A-Z]|[a-z])+)(?<row>(\d)+)")]
	private static partial Regex GetCellReferenceRegex();
}
