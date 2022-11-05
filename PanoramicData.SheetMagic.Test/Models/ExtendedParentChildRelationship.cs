namespace PanoramicData.SheetMagic.Test.Models;

public class ExtendedParentChildRelationship
{
	public int ParentDid { get; set; }

	public int RootDid { get; set; }

	public int ComponentDid { get; set; }

	public int MergeDid { get; set; }

	public string? Host { get; set; }

	public string? MissingThing { get; set; }
}