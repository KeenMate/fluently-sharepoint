using System.Xml.Serialization;

namespace TestConsole
{
	[XmlRoot(ElementName = "Stage", Namespace = "http://schemas.microsoft.com/office/infopath/2009/WSSList/dataFields")]
	public class Stage
	{
		[XmlAttribute(AttributeName = "nil", Namespace = "http://www.w3.org/2001/XMLSchema-instance")]
		public string Nil { get; set; }
	}
}