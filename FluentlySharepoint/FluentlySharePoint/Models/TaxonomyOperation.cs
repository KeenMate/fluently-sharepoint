using System;
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
		private readonly CSOMOperation operation;

		public TaxonomyOperation(CSOMOperation operation)
		{
			this.operation = operation;
		}

		[Obsolete("Not used yet")]
		public Queue<DeferredAction> ActionQueue { get; } = new Queue<DeferredAction>(10);
		
		public Dictionary<string, TermGroup> LoadedTermGroups { get; } = new Dictionary<string, TermGroup>(5);
		public Dictionary<string, TermSet> LoadedTermSets { get; } = new Dictionary<string, TermSet>(5);
		public Dictionary<string, Term> LoadedTerms { get; } = new Dictionary<string, Term>(5);

		public LockLevel LevelLock { get; set; }
		public TaxonomySession Session { get; set; }
		public TermStore LastTermStore { get; set; }

		public TermGroup LastTermGroup { get; set; }
		public TermSet LastTermSet { get; set; }
		public Term LastTerm { get; set; }


		[Obsolete("Not used yet")]
		private void ProcessLoaded(ClientObject clientObject)
		{
			switch (clientObject)
			{
				case TermGroup g:
					LoadedTermGroups.AddOrUpdate(g.Name.ToLower(), g);
					break;
				case TermSet s:
					LoadedTermSets.AddOrUpdate(s.Name.ToLower(), s);
					break;
				//case List t:
				//	LoadedLists[l.Title] = l;
				//	break;
				//case WebCollection wc:
				//	wc.ToList().ForEach(ProcessLoaded);
				//	break;
				//case ListCollection lc:
				//	lc.ToList().ForEach(ProcessLoaded);
				//	break;
			}
		}
	}
}