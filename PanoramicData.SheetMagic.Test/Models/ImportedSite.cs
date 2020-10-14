using System.ComponentModel;

namespace PanoramicData.SheetMagic.Test.Models
{
	public class ImportedSite
	{
		[Description("Estate Name")]
		public string EstateName { get; set; } = string.Empty;

		//public string AreaName { get; set; } = string.Empty;

		[Description("Area Name Hierarchy")]
		public string AreaNameHierarchy { get; set; } = string.Empty;

		//public string AreaParentName { get; set; } = string.Empty;

		[Description("Building Name")]
		public string BuildingName { get; set; } = string.Empty;

		[Description("Building Address")]
		public string BuildingAddress { get; set; } = string.Empty;

		//public string SchoolAddress { get; set; } = string.Empty;

		[Description("Building Latitude")]
		public double? BuildingLatitude { get; set; }

		[Description("Building Longitude")]
		public double? BuildingLongitude { get; set; }

		[Description("Floor Name")]
		public string FloorName { get; set; } = string.Empty;

		[Description("Floor RF Model")]
		public string FloorRfModel { get; set; } = string.Empty;

		[Description("Floor Width (Feet)")]
		public double FloorWidthFeet { get; set; }

		[Description("Floor Length (Feet)")]
		public double FloorLengthFeet { get; set; }

		[Description("Floor Height (Feet)")]
		public double FloorHeightFeet { get; set; }

		[Description("Network Tags")]
		public string NetworkTags { get; set; } = string.Empty;
	}
}
