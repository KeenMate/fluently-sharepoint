using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using KeenMate.FluentlySharePoint.Enums;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;

namespace KeenMate.FluentlySharePoint.Extensions
{
	public static class Site
	{
		public static WebTemplateCollection GetWebTemplates(this CSOMOperation operation, Lcid lcid,
			bool overrideCompatibilityLevel = false, Action<ClientContext, WebTemplateCollection> templatesLoader = null)
		{
			return GetWebTemplates(operation, (uint)lcid, overrideCompatibilityLevel, templatesLoader);
		}

		public static WebTemplateCollection GetWebTemplates(this CSOMOperation operation, uint lcid, bool overrideCompatibilityLevel = false, Action<ClientContext, WebTemplateCollection> templatesLoader = null)
		{
			operation.LogInfo($"Getting web templates for LCID {lcid}");

			var collection = operation.LastSite.GetWebTemplates(lcid, overrideCompatibilityLevel ? 1 : 0);

			if (templatesLoader != null)
			{
				templatesLoader(operation.Context, collection);
			}
			else
			{
				operation.Context.Load(collection);
			}

			operation.Execute();

			return collection;
		}

		public static TermCollection GetAllTerms(this CSOMOperation operation, string termSetName, Lcid lcid,
			params Expression<Func<Term, object>>[] keysToLoad)
		{
			return operation.GetAllTerms(termSetName, (uint) lcid, keysToLoad);
		}

		public static TermCollection GetAllTerms(this CSOMOperation operation, string termSetName, uint lcid, params Expression<Func<Term, object>>[] keysToLoad)
		{
			var termSets = TaxonomySession.GetTaxonomySession(operation.Context).GetTermSetsByName(termSetName, (int) lcid);

			operation.Context.Load(termSets);
			operation.Execute();

			if (termSets.Count == 0)
				throw new KeyNotFoundException($"Term set {termSetName} not found");

			var terms = termSets.First().GetAllTerms();

			operation.Context.Load(terms, t => t.Include(keysToLoad));
			operation.Context.ExecuteQuery();

			return terms;
		}
	}
}