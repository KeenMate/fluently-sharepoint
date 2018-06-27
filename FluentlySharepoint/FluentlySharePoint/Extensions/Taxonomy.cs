using System;
using Microsoft.SharePoint.Client.Taxonomy;

namespace KeenMate.FluentlySharePoint.Extensions
{
	public static class Taxonomy
	{
		public static CSOMOperation OpenTaxonomySession(this CSOMOperation operation)
		{
			operation.TaxonomySession = TaxonomySession.GetTaxonomySession(operation.Context);
			operation.Context.Load(operation.TaxonomySession);

			return operation;
		}

		public static CSOMOperation SelectTaxonomyStore(this CSOMOperation operation, string storeName="", Guid? storeGuid = null)
		{
			if (!string.IsNullOrEmpty(storeName))
				operation.TaxonomyStore = operation.TaxonomySession.TermStores.GetByName(storeName);
			else if (storeGuid.HasValue)
				operation.TaxonomyStore = operation.TaxonomySession.TermStores.GetById(storeGuid.Value);
			else
				operation.TaxonomyStore = operation.TaxonomySession.GetDefaultSiteCollectionTermStore();

			operation.Context.Load(operation.TaxonomyStore);

			return operation;
		}

		public static CSOMOperation CreateTaxonomyGroup(this CSOMOperation operation, string name)
		{
			operation.TaxonomyStore.CreateGroup()
		}
		
	}
}