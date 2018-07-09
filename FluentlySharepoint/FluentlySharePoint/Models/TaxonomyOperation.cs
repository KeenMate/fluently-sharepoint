using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;

namespace KeenMate.FluentlySharePoint.Models.Taxonomy
{
	public enum LockLevel
	{
		Undefined,
		TermSet,
		Term
	}
	public class TaxonomyOperation
	{
		public Dictionary<string, TermGroup> LoadedTermGroups { get; } = new Dictionary<string, TermGroup>(5);
		public Dictionary<string, TermSet> LoadedTermSets { get; } = new Dictionary<string, TermSet>(5);
		public Dictionary<string, Term> LoadedTerms { get; } = new Dictionary<string, Term>(5);

		public LockLevel LevelLock { get; set; }
		public TaxonomySession Session { get; set; }
		public TermStore LastTermStore { get; set; }

		public TermGroup LastTermGroup { get; set; }
		public TermSet LastTermSet { get; set; }
		public Term LastTerm { get; set; }
		
	}
}