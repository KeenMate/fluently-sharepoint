using System;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using Microsoft.SharePoint.Client;

namespace KeenMate.FluentlySharePoint.Models
{
	[XmlRoot("Field")]
	public class FieldCreationInformation
	{
		[XmlAttribute("ID")]
		public Guid Id { get; set; }

		[XmlAttribute]
		public string DisplayName { get; set; }

		[XmlAttribute("Name")]
		public string InternalName { get; set; }

		[XmlIgnore]
		public bool AddToDefaultView { get; set; }


		//public IEnumerable<KeyValuePair<string, string>> AdditionalAttributes { get; set; }

		[XmlAttribute("Type")]
		public FieldType FieldType { get; set; }

		[XmlAttribute]
		public string Group { get; set; }

		[XmlAttribute]
		public bool Required { get; set; }

		[XmlAttribute("EnforceUniqueValues")]
		public bool UniqueValues { get; set; }

		public string ToXml()
		{
			var serializer = new XmlSerializer(GetType());
			var settings = new XmlWriterSettings();
			settings.Indent = true;
			settings.OmitXmlDeclaration = true;
			var emptyNamepsaces = new XmlSerializerNamespaces(new[] { XmlQualifiedName.Empty });

			using (var stream = new StringWriter())
			using (var writer = XmlWriter.Create(stream, settings))
			{
				serializer.Serialize(writer, this, emptyNamepsaces);
				return stream.ToString();
			}
		}

		public FieldCreationInformation()
		{
			Id = Guid.NewGuid();
		}

	}
}