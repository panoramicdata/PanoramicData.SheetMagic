using System.Globalization;

namespace PanoramicData.SheetMagic;

/// <summary>
/// Cell creation and manipulation methods
/// </summary>
public partial class MagicSpreadsheet
{
	private static Cell CreateCell(string header, uint index, object? @object)
	{
		if (@object == null)
		{
			return CreateTextCell(header, index, string.Empty);
		}

		var objectTypeName = @object.GetType().Name;

		return objectTypeName switch
		{
			nameof(Int16) or nameof(Int32) or nameof(Int64) or
			nameof(UInt16) or nameof(UInt32) or nameof(UInt64)
				=> CreateNumericCell(header, index, Convert.ToDouble(@object)),

			nameof(Single) or nameof(Double) or nameof(Decimal)
				=> CreateNumericOrSpecialCell(header, index, @object),

			nameof(Boolean)
				=> CreateBooleanCell(header, index, (bool)@object),

			nameof(DateTime)
				=> CreateDateCell(header, index, (DateTime)@object),

			nameof(DateTimeOffset)
				=> CreateDateCell(header, index, ((DateTimeOffset)@object).UtcDateTime),

			_ when objectTypeName.StartsWith("Nullable`1")
				=> CreateNullableTypeCell(header, index, @object),

			_ => CreateTextCell(header, index, @object.ToString() ?? string.Empty),
		};
	}

	private static Cell CreateNumericOrSpecialCell(string header, uint index, object @object)
	{
		var doubleValue = Convert.ToDouble(@object);

		if (double.IsNaN(doubleValue))
		{
			return CreateTextCell(header, index, string.Empty);
		}

		if (double.IsPositiveInfinity(doubleValue))
		{
			return CreateTextCell(header, index, "Infinity");
		}

		if (double.IsNegativeInfinity(doubleValue))
		{
			return CreateTextCell(header, index, "-Infinity");
		}

		return CreateNumericCell(header, index, doubleValue);
	}

	private static Cell CreateNullableTypeCell(string header, uint index, object @object)
	{
		var objectTypeName = @object.GetType().Name;

		return objectTypeName switch
		{
			"Nullable`1<Single>" or "Nullable`1<Double>" or "Nullable`1<Decimal>"
				=> CreateNumericOrSpecialCell(header, index, @object),

			"Nullable`1<Boolean>"
				=> CreateBooleanCell(header, index, (bool)@object),

			"Nullable`1<DateTime>"
				=> CreateDateCell(header, index, (DateTime)@object),

			"Nullable`1<DateTimeOffset>"
				=> CreateDateCell(header, index, ((DateTimeOffset)@object).UtcDateTime),

			_ when objectTypeName.Contains("Int")
				=> CreateNumericCell(header, index, Convert.ToDouble(@object)),

			_ => CreateTextCell(header, index, @object.ToString() ?? string.Empty),
		};
	}

	private static Cell CreateNumericCell(string header, uint index, double number) =>
		new(new CellValue(number.ToString(CultureInfo.InvariantCulture)))
		{
			DataType = CellValues.Number,
			CellReference = header + index
		};

	private static Cell CreateDateCell(string header, uint index, DateTime date) =>
		new(new CellValue(date))
		{
			DataType = CellValues.Date,
			CellReference = header + index,
			StyleIndex = 1
		};

	private static Cell CreateTextCell(string header, uint index, string text) =>
		 new(new InlineString(new Text { Text = text }))
		 {
			 DataType = CellValues.InlineString,
			 CellReference = header + index
		 };

	private static Cell CreateBooleanCell(string header, uint index, bool booleanValue) =>
		new(new CellValue(booleanValue))
		{
			DataType = CellValues.Boolean,
			CellReference = header + index
		};
}
