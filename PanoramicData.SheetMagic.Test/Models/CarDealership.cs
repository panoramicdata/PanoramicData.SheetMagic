using System.Collections.Generic;

namespace PanoramicData.SheetMagic.Test.Models;

public class CarDealership
{
	public string Name { get; set; } = "Unnamed";
	public List<Car?> Cars { get; set; } = new List<Car?>();
}
