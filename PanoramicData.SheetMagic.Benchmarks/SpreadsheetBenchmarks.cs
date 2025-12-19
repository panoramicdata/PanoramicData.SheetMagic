using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Exporters;
using BenchmarkDotNet.Exporters.Csv;

namespace PanoramicData.SheetMagic.Benchmarks;

/// <summary>
/// Benchmarks for MagicSpreadsheet operations.
/// Run with: dotnet run -c Release
/// Results are saved to BenchmarkDotNet.Artifacts folder with HTML charts.
/// </summary>
[MemoryDiagnoser]
[SimpleJob(warmupCount: 3, iterationCount: 10)]
[HtmlExporter]
[CsvExporter]
[RPlotExporter]
[MarkdownExporter]
public class SpreadsheetBenchmarks
{
	private List<BenchmarkItem> _smallDataset = null!;
	private List<BenchmarkItem> _mediumDataset = null!;
	private List<BenchmarkItem> _largeDataset = null!;

	[GlobalSetup]
	public void Setup()
	{
		_smallDataset = GenerateData(100);
		_mediumDataset = GenerateData(1000);
		_largeDataset = GenerateData(10000);
	}

	private static List<BenchmarkItem> GenerateData(int count)
	{
		var items = new List<BenchmarkItem>(count);
		for (var i = 0; i < count; i++)
		{
			items.Add(new BenchmarkItem
			{
				Id = i,
				Name = $"Item {i}",
				Description = $"Description for item {i} with some additional text to make it longer",
				Value = i * 1.5,
				CreatedDate = DateTime.Now.AddDays(-i),
				IsActive = i % 2 == 0,
				Category = $"Category {i % 10}"
			});
		}
		return items;
	}

	[Benchmark(Description = "Write 100 rows")]
	public void WriteSmallDataset()
	{
		using var stream = new MemoryStream();
		using var spreadsheet = new MagicSpreadsheet(stream);
		spreadsheet.AddSheet(_smallDataset, "Data");
		spreadsheet.Save();
	}

	[Benchmark(Description = "Write 1,000 rows")]
	public void WriteMediumDataset()
	{
		using var stream = new MemoryStream();
		using var spreadsheet = new MagicSpreadsheet(stream);
		spreadsheet.AddSheet(_mediumDataset, "Data");
		spreadsheet.Save();
	}

	[Benchmark(Description = "Write 10,000 rows")]
	public void WriteLargeDataset()
	{
		using var stream = new MemoryStream();
		using var spreadsheet = new MagicSpreadsheet(stream);
		spreadsheet.AddSheet(_largeDataset, "Data");
		spreadsheet.Save();
	}

	[Benchmark(Description = "Write and read 100 rows")]
	public List<BenchmarkItem?> WriteAndReadSmall()
	{
		using var stream = new MemoryStream();
		
		using (var writeSpreadsheet = new MagicSpreadsheet(stream))
		{
			writeSpreadsheet.AddSheet(_smallDataset, "Data");
			writeSpreadsheet.Save();
		}

		stream.Position = 0;
		
		using var readSpreadsheet = new MagicSpreadsheet(stream);
		readSpreadsheet.Load();
		return readSpreadsheet.GetList<BenchmarkItem>("Data");
	}

	[Benchmark(Description = "Write and read 1,000 rows")]
	public List<BenchmarkItem?> WriteAndReadMedium()
	{
		using var stream = new MemoryStream();
		
		using (var writeSpreadsheet = new MagicSpreadsheet(stream))
		{
			writeSpreadsheet.AddSheet(_mediumDataset, "Data");
			writeSpreadsheet.Save();
		}

		stream.Position = 0;
		
		using var readSpreadsheet = new MagicSpreadsheet(stream);
		readSpreadsheet.Load();
		return readSpreadsheet.GetList<BenchmarkItem>("Data");
	}

	[Benchmark(Description = "Write multiple sheets")]
	public void WriteMultipleSheets()
	{
		using var stream = new MemoryStream();
		using var spreadsheet = new MagicSpreadsheet(stream);
		
		for (var i = 0; i < 5; i++)
		{
			spreadsheet.AddSheet(_smallDataset, $"Sheet{i}");
		}
		
		spreadsheet.Save();
	}

	[Benchmark(Description = "Write with no table style")]
	public void WriteNoTableStyle()
	{
		using var stream = new MemoryStream();
		using var spreadsheet = new MagicSpreadsheet(stream);
		spreadsheet.AddSheet(_mediumDataset, "Data", new AddSheetOptions { TableOptions = null });
		spreadsheet.Save();
	}
}

/// <summary>
/// Sample data class for benchmarking.
/// </summary>
public class BenchmarkItem
{
	public int Id { get; set; }
	public string Name { get; set; } = string.Empty;
	public string Description { get; set; } = string.Empty;
	public double Value { get; set; }
	public DateTime CreatedDate { get; set; }
	public bool IsActive { get; set; }
	public string Category { get; set; } = string.Empty;
}
