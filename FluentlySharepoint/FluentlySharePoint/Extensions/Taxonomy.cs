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

		
	}
}