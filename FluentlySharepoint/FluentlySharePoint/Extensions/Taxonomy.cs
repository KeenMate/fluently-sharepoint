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
			operation.Execute();

			return operation;
		}

		public static CSOMOperation SelectTaxonomyStore(this CSOMOperation operation, string storeName="", Guid storeGuid = new Guid())
		{
			if (!string.IsNullOrEmpty(storeName))
				operation.TaxonomyStore = operation.TaxonomySession.TermStores.GetByName(storeName);
			else if (storeGuid != Guid.Empty)
				operation.TaxonomyStore = operation.TaxonomySession.TermStores.GetById(storeGuid);
			else
				operation.TaxonomyStore = operation.TaxonomySession.GetDefaultSiteCollectionTermStore();

			operation.Context.Load(operation.TaxonomyStore);

			return operation;
		}



	}
}