using System.ComponentModel;

namespace PanoramicData.SheetMagic.Test.Models
{
	internal class FunkyAnimal
	{
		[Description("Leg Count")]
		public int Leg_Count { get; set; }

		public int Id { get; set; }

		public string? Name { get; set; }

		[Description("Weight KG")]
		public double WeightKg { get; set; }

		public string? Description { get; set; }
	}
}