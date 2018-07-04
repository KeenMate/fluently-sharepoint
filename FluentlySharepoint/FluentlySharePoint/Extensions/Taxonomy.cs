using System;
using System.Collections.Generic;
using KeenMate.FluentlySharePoint.Models;
using Microsoft.SharePoint.Client.Taxonomy;

namespace KeenMate.FluentlySharePoint.Extensions
{
	public static class Taxonomy
	{
		private static void EnsureTaxonomyOperation(CSOMOperation operation)
		{
			operation.TaxonomyOperation = operation.TaxonomyOperation ?? new TaxonomyOperation();
		}
		public static CSOMOperation OpenTaxonomySession(this CSOMOperation operation)
		{
			operation.TaxonomyOperation.Session = TaxonomySession.GetTaxonomySession(operation.Context);
			operation.Context.Load(operation.TaxonomyOperation.Session);

			return operation;
		}

		public static CSOMOperation SelectTaxonomyStore(this CSOMOperation operation, string storeName="", Guid? storeGuid = null)
		{
			var op = operation.TaxonomyOperation;
			if (!string.IsNullOrEmpty(storeName))
				op.LastTermStore = op.Session.TermStores.GetByName(storeName);
			else if (storeGuid.HasValue)
				op.LastTermStore = op.Session.TermStores.GetById(storeGuid.Value);
			else
				op.LastTermStore = op.Session.GetDefaultSiteCollectionTermStore();

			operation.Context.Load(op.LastTermStore);

			return operation;
		}

		public static CSOMOperation CreateTermGroup(this CSOMOperation operation, string name, Guid? guid = null)
		{
			var op = operation.TaxonomyOperation;

			op.LastTermGroup = op.LastTermStore.CreateGroup(name, guid ?? Guid.NewGuid());

			return operation;
		}

		public static CSOMOperation CreateTermSet(this CSOMOperation operation, string name, Guid? guid = null, int lcid = 1033)
		{
			var op = operation.TaxonomyOperation;

			op.LastTermSet = op.LastTermGroup.CreateTermSet(name, guid??Guid.NewGuid(), lcid);

			return operation;
		}

		public static CSOMOperation CreateTerm(this CSOMOperation operation, string name
			, Guid? guid = null, int lcid = 1033, IEnumerable<Tuple<string, int>> descriptions = null)
		{
			var op = operation.TaxonomyOperation;

			op.LastTerm = op.LastTermSet.CreateTerm(name, lcid, guid ?? Guid.NewGuid());

			if (descriptions != null)
				foreach (var description in descriptions)
				{
					op.LastTerm.SetDescription(description.Item1, description.Item2);
				}

			return operation;
		}

	}
}