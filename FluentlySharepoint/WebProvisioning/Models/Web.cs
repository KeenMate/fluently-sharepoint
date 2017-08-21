using System.Collections.Generic;

namespace WebProvisioning.Models
{
	public class Web
	{
		public string Name { get; set; }
		public string Url { get; set; }
		public string Template { get; set; }
		public List<List> Lists { get; set; }
	}
}