using System;
using System.Linq.Expressions;
using KeenMate.FluentlySharePoint.Enums;
using KeenMate.FluentlySharePoint.Helpers;
using Microsoft.SharePoint.Client;

namespace KeenMate.FluentlySharePoint.Extensions
{
	public static class WebExtensions
	{
		public static CSOMOperation LoadWebs(this CSOMOperation operation, params Expression<Func<Microsoft.SharePoint.Client.Web, object>>[] keysToLoad) // todo add custom loader
		{
			operation.LogDebug("Loading all webs");

			var webs = operation.DecideWeb().Webs;

			operation.Context.Load(webs, CSOMOperation.DefaultRetrievals.WebCollection);

			if (keysToLoad.Length > 0)
				operation.Context.Load(webs, w => w.Include(keysToLoad));

			operation.ActionQueue.Enqueue(new DeferredAction { ClientObject = webs, Action = DeferredActions.Load });

			return operation;
		}

		public static CSOMOperation LoadWeb(this CSOMOperation operation, Guid id,
			params Expression<Func<Web, object>>[] retrievals)
		{
			operation.LogInfo($"Loading web with id: {id}");

			Web web = operation.LastSite.OpenWebById(id);

			return LoadWeb(operation, web, retrievals);
		}

		public static CSOMOperation LoadWeb(this CSOMOperation operation, string name,
			params Expression<Func<Web, object>>[] retrievals)
		{
			operation.LogDebug($"Loading web");

			var web = operation.LastSite.OpenWeb(name);

			return LoadWeb(operation, web, retrievals);
		}

		public static CSOMOperation LoadWebByUrl(this CSOMOperation operation, string url,
			params Expression<Func<Web, object>>[] retrievals)
		{
			operation.LogDebug($"Loading web with url: {url}");

			var web = operation.LastSite.OpenWebUsingPath(ResourcePath.FromDecodedUrl(url));

			return LoadWeb(operation, web, retrievals);
		}

		public static CSOMOperation LoadWeb(this CSOMOperation operation, Web web,
			params Expression<Func<Web, object>>[] retrievals)
		{
			operation.LogDebug($"Loading web");

			operation.Context.Load(web,
				retrievals != null && retrievals.Length > 0
					? retrievals
					: CSOMOperation.DefaultRetrievals.Web);

			operation.SetLevel(OperationLevels.Web, web);
			operation.ActionQueue.Enqueue(new DeferredAction { ClientObject = web, Action = DeferredActions.Load });

			return operation;
		}

		public static CSOMOperation GetWebTemplates(this CSOMOperation operation, uint? lcid = null, int? compatibilityLevel = null)
		{
			lcid = lcid ?? operation.DefaultLcid;
			compatibilityLevel = compatibilityLevel ?? operation.DefaultCompatibilityLevel;

			var templates = operation.LastSite.GetWebTemplates(lcid.Value, compatibilityLevel.Value);

			operation.Load(templates, CSOMOperation.DefaultRetrievals.WebTemplateCollection);
			return operation;
		}

		public static ListTemplateCollection GetListTemplates(this CSOMOperation operation, params Expression<Func<ListTemplateCollection, object>>[] retrievals)
		{
			var templates = operation.LastSite.GetCustomListTemplates(operation.DecideWeb());

			operation.Context.Load(templates,
				retrievals != null && retrievals.Length > 0
					? retrievals
					: CSOMOperation.DefaultRetrievals.ListTemplateCollection);

			operation.Execute();

			return templates;
		}

		public static CSOMOperation SelectWeb(this CSOMOperation operation, string url)
		{
			var key = url.ToLower();
			operation.LogDebug($"Selecting web with url: {key}");
			if (operation.LoadedWebs.ContainsKey(key))
			{
				operation.SetLevel(OperationLevels.Web, operation.LoadedWebs[key]);
			}
			else
			{
				throw new ArgumentException($"Web with URL {key} doesn't exists");
			}

			return operation;
		}

		public static CSOMOperation CreateWeb(this CSOMOperation operation, string title, int? lcid, string url = "", string template = "")
		{
			operation.LogInfo($"Creating web {title}");

			url = url.IsNotNullOrEmpty() ? url : operation.NormalizeUrl(title);
			Web rootWeb = operation.DecideWeb();

			lcid = (int)((uint?)lcid ?? operation.DecideWeb().Language);

			operation.LogDebug($"Web creation information set to Title: {title}, Url: {url}, Lcid: {lcid}, Template: {template}");
			WebCreationInformation webInformation = new WebCreationInformation
			{
				Title = title,
				Url = url,
				WebTemplate = template,
				Language = lcid.Value
			};

			var web = rootWeb.Webs.Add(webInformation);

			operation.LoadWebWithDefaultRetrievals(web);

			operation.SetLevel(OperationLevels.Web, web);
			operation.ActionQueue.Enqueue(new DeferredAction { ClientObject = web, Action = DeferredActions.Load });

			return operation;
		}

		public static CSOMOperation DeleteWeb(this CSOMOperation operation)
		{
			operation.LogInfo($"Deleting last selected web: {operation.LastWeb.Title}");

			operation.ActionQueue.Enqueue(new DeferredAction { ClientObject = operation.LastWeb, Action = DeferredActions.Delete });

			return operation;
		}

		public static Web DecideWeb(this CSOMOperation operation)
		{
			operation.LogTrace($"Deciding which web to use by operation level: {operation.OperationLevel}");
			switch (operation.OperationLevel)
			{
				case OperationLevels.Site:
					operation.LogTrace("Using last site root web");
					return operation.LastSite.RootWeb;
				case OperationLevels.Web:
					operation.LogTrace("Using last loaded web");
					return operation.LastWeb;
				case OperationLevels.List:
					operation.LogTrace("Using list's parent web");
					return operation.LastList.ParentWeb;
				default:
					operation.LogWarn("Not sure how did you end up here but for the current operation level there is no predefined web to load.");
					return null; //todo throw exception
			}
		}
	}
}