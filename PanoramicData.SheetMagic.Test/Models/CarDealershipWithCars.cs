using System.Collections.Generic;

namespace PanoramicData.SheetMagic.Test.Models;

public class CarDealershipWithCars : CarDealership
{
	public List<Car?> Cars { get; init; } = [];
}
