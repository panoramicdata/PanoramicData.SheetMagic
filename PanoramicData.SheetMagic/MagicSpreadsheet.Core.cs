using Sheet = DocumentFormat.OpenXml.Spreadsheet.Sheet;

namespace PanoramicData.SheetMagic;

/// <summary>
/// Core functionality for MagicSpreadsheet - constructors, fields, properties, and lifecycle methods.
/// Provides easy saving and loading of generic lists to/from XLSX workbooks.
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

	/// <summary>
	/// Creates a new MagicSpreadsheet instance for the specified file with options.
	/// </summary>
	/// <param name="fileInfo">The file to read from or write to.</param>
	/// <param name="options">Configuration options.</param>
	public MagicSpreadsheet(FileInfo fileInfo, Options options)
	{
		_fileInfo = fileInfo;
		_options = options;
	}

	/// <summary>
	/// Creates a new MagicSpreadsheet instance for the specified file with default options.
	/// </summary>
	/// <param name="fileInfo">The file to read from or write to.</param>
	public MagicSpreadsheet(FileInfo fileInfo)
		: this(fileInfo, new Options())
	{
	}

	/// <summary>
	/// Creates a new MagicSpreadsheet instance for the specified stream with options.
	/// </summary>
	/// <param name="stream">The stream to read from or write to.</param>
	/// <param name="options">Configuration options.</param>
	public MagicSpreadsheet(Stream stream, Options options)
	{
		_stream = stream;
		_options = options;
	}

	/// <summary>
	/// Creates a new MagicSpreadsheet instance for the specified stream with default options.
	/// </summary>
	/// <param name="stream">The stream to read from or write to.</param>
	public MagicSpreadsheet(Stream stream)
		: this(stream, new Options())
	{
	}

	/// <summary>
	/// Gets the names of all sheets in the loaded workbook.
	/// </summary>
	/// <exception cref="InvalidOperationException">Thrown if no document is loaded.</exception>
	public List<string> SheetNames
		=> [.. ((((_document ?? throw new InvalidOperationException("No document loaded."))
			.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart not created"))
			.Workbook ?? throw new InvalidOperationException("Workbook not created"))
			.Sheets ?? throw new InvalidOperationException("Sheets not created"))
			.ChildElements
			.Cast<Sheet>()
			.Select(static s => s.Name?.Value ?? string.Empty)];

	/// <summary>
	/// Loads the spreadsheet document for reading.
	/// </summary>
	public void Load() => _document = _fileInfo is not null
		? SpreadsheetDocument.Open(_fileInfo.FullName, false)
		: SpreadsheetDocument.Open(_stream!, false);

	/// <summary>
	/// Saves the spreadsheet document to the file or stream.
	/// </summary>
	/// <exception cref="InvalidOperationException">Thrown if the document was not created correctly.</exception>
	public void Save()
	{
		// Ensure that at least one sheet has been added
		if (_document?.WorkbookPart?.Workbook?.Sheets == null || !_document.WorkbookPart.Workbook.Sheets.Any())
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

	/// <summary>
	/// Disposes of the spreadsheet resources.
	/// </summary>
	public void Dispose()
	{
		ReleaseUnmanagedResources();
		GC.SuppressFinalize(this);
	}

	/// <summary>
	/// Finalizer to ensure resources are released.
	/// </summary>
	~MagicSpreadsheet()
	{
		ReleaseUnmanagedResources();
	}

	[GeneratedRegex(@"(?<col>([A-Z]|[a-z])+)(?<row>(\d)+)")]
	private static partial Regex GetCellReferenceRegex();
}
