namespace PanoramicData.SheetMagic.Test.Models;

public class DeviceSpecification
{
	public string? HostName { get; set; }

	public string? DeviceDisplayName { get; set; }

	public string? DeviceDescription { get; set; }

	public string? DeviceGroups { get; set; }

	public string? PreferredCollector { get; set; }

	public bool? EnableAlerts { get; set; }

	public bool? EnableNetflow { get; set; }

	public string? NetflowCollector { get; set; }

	public string? Link { get; set; }
}
