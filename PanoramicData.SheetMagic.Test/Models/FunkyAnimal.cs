using System.Collections.Generic;
using System.ComponentModel;

namespace PanoramicData.SheetMagic.Test.Models;

internal class FunkyAnimal
{
	[Description("Leg Count")]
	public int Leg_Count { get; set; }

	public int Id { get; set; }

	public string? Name { get; set; }

	[Description("Weight KG")]
	public double WeightKg { get; set; }

	public string? Description { get; set; }

	public List<string>? Nicknames { get; set; }

	public List<FunkyAnimal> Friends { get; set; } = new();
}