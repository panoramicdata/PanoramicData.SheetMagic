using System.ComponentModel;

namespace PanoramicData.SheetMagic.Test.Models
{
	public class ImportedDevice
	{
		[Description("Estate Name")]
		public string EstateName { get; set; } = string.Empty;

		[Description("Site Name")]
		public string SiteName { get; set; } = string.Empty;

		[Description("Legacy Device Serial Number")]
		public string LegacyDeviceSerialNumber { get; set; } = string.Empty;

		[Description("Serial Number")]
		public string NewDeviceSerialNumber { get; set; } = string.Empty;

		[Description("Stack Top Serial Number")]
		public string? StackTopSerialNumber { get; set; }

		[Description("Stack Position")]
		public int? StackPosition { get; set; }

		[Description("Hostname")]
		public string? Hostname { get; set; }

		[Description("Management VLAN")]
		public int? ManagementVlan { get; set; }

		[Description("Management IP Address")]
		public string? ManagementIpAddress { get; set; }

		[Description("Onboarding Configuration Template Pre")]
		public string? OnboardingConfigurationTemplatePreConfigurationTemplateName { get; set; }

		[Description("Onboarding Configuration Template Post")]
		public string? OnboardingConfigurationTemplatePostConfigurationTemplateName { get; set; }

		[Description("Switch Profile")]
		public string? SwitchProfile { get; set; }

		[Description("Rf Profile")]
		public string? RfProfile { get; set; }

		[Description("Image")]
		public string? ImageName { get; set; }

		[Description("Network Module")]
		public string NewDeviceNetworkModule { get; set; } = string.Empty;

		[Description("Notes")]
		public string Notes { get; set; } = string.Empty;
	}
}
