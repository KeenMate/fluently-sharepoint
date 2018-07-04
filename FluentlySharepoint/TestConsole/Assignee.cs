using System.Collections.Generic;
using System.Xml.Serialization;

namespace TestConsole
{
	[XmlRoot(ElementName = "Assignee", Namespace = "http://schemas.microsoft.com/office/infopath/2009/WSSList/dataFields")]
	public class Assignee
	{
		[XmlElement(ElementName = "Person", Namespace = "http://schemas.microsoft.com/office/infopath/2007/PartnerControls")]
		public List<Person> Person { get; set; }
	}
}