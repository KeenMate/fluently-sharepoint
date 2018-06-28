using Microsoft.SharePoint.Client.Taxonomy;

namespace KeenMate.FluentlySharePoint.Models
{
	public class TaxonomyOperation
	{
		public TaxonomySession Session { get; set; }
		public TermStore LastTermStore { get; set; }

		public TermGroup LastTermGroup { get; set; }
		public TermSet LastTermSet { get; set; }
		public Term LastTerm { get; set; }
		
	}
}