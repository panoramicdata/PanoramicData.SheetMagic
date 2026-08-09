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
	private bool _isSaved;

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
	/// <exception cref="SpreadsheetTooLargeException">
	/// Thrown if the serialised workbook exceeds the maximum supported size.
	/// </exception>
	public void Save()
	{
		if (_isSaved)
		{
			throw new InvalidOperationException("The spreadsheet has already been saved and cannot be saved again.");
		}

		_isSaved = true;

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

		var document = _document;

		// Once a save has been attempted the document is no longer usable, whether or not the
		// attempt succeeded.  Drop the reference before doing the work so that a subsequent
		// Dispose() cannot retry the save and throw a second time.
		_document = null;

		try
		{
			document.WorkbookPart.Workbook.Save();

			// Disposing the package is what commits the remaining parts and closes the file.
			// It is done here, rather than in Dispose(), so that any failure surfaces from the
			// caller's Save() call where it can be caught.
			document.Dispose();
		}
		catch (Exception e) when (IsSizeLimitException(e))
		{
			throw new SpreadsheetTooLargeException(e);
		}

		// Do we have a stream?
		if (_stream is not null)
		{
			// YES - Ensure it's flushed and seek back to the beginning for consumption
			_stream.Flush();
			_ = _stream.Seek(0, SeekOrigin.Begin);
		}
	}

	/// <summary>
	/// Determines whether an exception raised while serialising indicates that the workbook
	/// exceeded the maximum stream length.
	/// </summary>
	/// <remarks>
	/// The OpenXML writer reports this as a plain <see cref="IOException"/> whose only
	/// distinguishing feature is its message, so message inspection is the only option.  Any
	/// other failure (disk full, file locked, and so on) is deliberately left untranslated.
	/// </remarks>
	private static bool IsSizeLimitException(Exception e)
		=> e is IOException or NotSupportedException
			&& e.Message.Contains("too long", StringComparison.OrdinalIgnoreCase);

	/// <summary>
	/// Disposes of the spreadsheet resources.
	/// </summary>
	/// <remarks>
	/// Safe to call more than once, and safe to call after <see cref="Save"/>.  There is
	/// deliberately no finalizer: <see cref="MagicSpreadsheet"/> owns no unmanaged resources
	/// directly, and disposing an <see cref="SpreadsheetDocument"/> can perform work that
	/// throws.  On the finalizer thread such an exception is unhandleable and terminates the
	/// process, so that path must not exist.
	/// </remarks>
	public void Dispose()
	{
		var document = _document;
		_document = null;
		document?.Dispose();
	}

	[GeneratedRegex(@"(?<col>([A-Z]|[a-z])+)(?<row>(\d)+)")]
	private static partial Regex GetCellReferenceRegex();
}
