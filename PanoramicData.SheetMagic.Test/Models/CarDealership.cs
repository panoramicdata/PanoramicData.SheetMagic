using System;

namespace PanoramicData.SheetMagic.Test.Models
{
	public class CarDealership
	{
		public string Name { get; init; } = "Unnamed";

		public DateTime Founded { get; init; }

		public bool IsPrivatelyOwned { get; init; }

		public int UkRanking { get; init; }

		public int? EmployeeCount { get; init; }

		public DateTimeOffset? ClosureDate { get; init; }
	}
}