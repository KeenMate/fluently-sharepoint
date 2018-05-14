using System.Xml.Serialization;

namespace KeenMate.FluentlySharePoint.Enums
{
	public enum ChoiceTypes
	{
		Default,
		[XmlEnum]
		Dropdown,
		[XmlEnum]
		RadioButtons
	}
}