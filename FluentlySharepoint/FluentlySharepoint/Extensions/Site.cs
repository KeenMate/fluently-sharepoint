using System;
using KeenMate.FluentlySharePoint.Enums;
using Microsoft.SharePoint.Client;

namespace KeenMate.FluentlySharePoint.Extensions
{
	public static class Site
	{
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

		public static WebTemplateCollection GetWebTemplates(this CSOMOperation operation, Lcid lcid,
			bool overrideCompatibilityLevel = false, Action<ClientContext, WebTemplateCollection> templatesLoader = null)
		{
			return GetWebTemplates(operation, (uint) lcid, overrideCompatibilityLevel, templatesLoader);
		}
	}
}