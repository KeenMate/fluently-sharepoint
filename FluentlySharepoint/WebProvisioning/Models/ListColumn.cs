using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace WebProvisioning.Models
{
	public class ListColumn
	{
		public string Name { get; set; }
		public string DisplayName { get; set; }
		[JsonConverter(typeof(StringEnumConverter))]
		public FieldType Type { get; set; }
		public bool Required { get; set; }
		public bool UniqueValues { get; set; }
	}
}