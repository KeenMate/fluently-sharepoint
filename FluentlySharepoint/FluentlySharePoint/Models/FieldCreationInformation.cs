using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;
using KeenMate.FluentlySharePoint.Enums;
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
		
		[XmlIgnore] public int? MaxLegth { get; set; }
		[XmlAttribute("MaxLength")]
		public int MaxLengthSerializable { get => MaxLegth ?? 0; set => MaxLegth = value; }
		public bool MaxLengthSerializableSpecified => MaxLegth.HasValue;

		[XmlAttribute]
		public ChoiceTypes Format { get; set; }

		public bool ChoiceTypeSpecified => Format != ChoiceTypes.Default;

		[XmlAttribute("EnforceUniqueValues")]
		public bool UniqueValues { get; set; }

		[XmlAttribute]
		public bool Percentage { get; set; }

		[XmlAttribute]
		public int Decimals { get; set; }

		[XmlAttribute]
		public string List { get; set; }

		[XmlAttribute]
		public string  ShowField { get; set; }
		public bool ShowFieldSpecified => FieldType == FieldType.Lookup;

		[XmlAttribute]
		public RelationshipDeleteBehaviorType RelationshipDeleteBehaviorType { get; set; }
		[XmlIgnore] public int? Min { get; set; }
		[XmlAttribute("Min")]
		public int MinSerializable { get => Min ?? 0; set => Min = value; }
		public bool MinSerializableSpecified => Min.HasValue;

		[XmlIgnore] public int? Max { get; set; }
		[XmlAttribute("Max")]
		public int MaxSerializable { get => Max ?? 0; set => Max = value; }
		public bool MaxSerializableSpecified => Max.HasValue;

		[XmlElement]
		public string Default { get; set; }

		[XmlArray("CHOICES")]
		[XmlArrayItem("CHOICE")]
		public List<string> Choices { get; set; }

		public string ToXml()
		{
			var serializer = new XmlSerializer(GetType());
			var settings = new XmlWriterSettings
			{
				Indent = true,
				OmitXmlDeclaration = true
			};
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