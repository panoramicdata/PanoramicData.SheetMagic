namespace PanoramicData.SheetMagic.Test.Models
{
	public class Car
	{
		public int WheelCount { get; set; }
		public int Id { get; set; }
		public string? Name { get; set; }
		public int WeightKg { get; set; }

		public override string ToString() => $"{Name ?? "Unnamed"} ({WeightKg}kg)";
	}
}