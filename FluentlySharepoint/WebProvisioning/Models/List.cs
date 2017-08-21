using System.Collections.Generic;

namespace WebProvisioning.Models
{
	public class List
	{
		public string Title { get; set; }
		public string Template { get; set; }
		public List<ListColumn> Columns { get; set; }
	}
}