using System;
using Microsoft.SharePoint.Client;

namespace KeenMate.FluentlySharePoint.Extensions
{
	public static class Web
	{
		public static CSOMOperation LoadWebs(this CSOMOperation operation) // todo add custom loader
		{
			var webs = operation.DecideWeb().Webs;

			operation.Context.Load(webs);
			operation.ActionQueue.Enqueue(new DeferredAction { ClientObject = webs, Action = DeferredActions.Load });

			return operation;
		}

		public static CSOMOperation LoadWeb(this CSOMOperation operation, string name = "",
			Action<ClientContext, Microsoft.SharePoint.Client.Web> webLoader = null)
		{
			var web = operation.LastSite.OpenWeb(name);

			if (webLoader != null)
				webLoader(operation.Context, operation.LastWeb);
			else
			{
				operation.Context.Load(operation.LastWeb);
			}

			operation.SetLevel(OperationLevels.Web, web);
			operation.ActionQueue.Enqueue(new DeferredAction { ClientObject = operation.LastWeb, Action = DeferredActions.Load });

			return operation;
		}

		public static CSOMOperation SelectWeb(this CSOMOperation operation, string url)
		{
			if (operation.LoadedWebs.ContainsKey(url))
			{
				operation.SetLevel(OperationLevels.Web, operation.LoadedWebs[url]);
			}
			else
			{
				throw new ArgumentException($"Web with URL {url} doesn't exists");
			}

			return operation;
		}

		public static CSOMOperation CreateWeb(this CSOMOperation operation, string name, string url = "", string template = "")
		{
			WebCreationInformation webInformation = new WebCreationInformation
			{
				Title = name,
				Url = String.IsNullOrEmpty(url) ? name : url,
				WebTemplate = template
			};

			var web = operation.DecideWeb().Webs.Add(webInformation);

			operation.Context.Load(web);
			operation.SetLevel(OperationLevels.Web, web);
			operation.ActionQueue.Enqueue(new DeferredAction { ClientObject = web, Action = DeferredActions.Load });

			return operation;
		}

		public static CSOMOperation DeleteWeb(this CSOMOperation operation)
		{
			operation.ActionQueue.Enqueue(new DeferredAction {ClientObject = operation.LastWeb, Action = DeferredActions.Delete});
			return operation;
		}

		public static Microsoft.SharePoint.Client.Web DecideWeb(this CSOMOperation operation)
		{
			switch (operation.OperationLevel)
			{
				case OperationLevels.Site:
					return operation.LastSite.RootWeb;
				case OperationLevels.Web:
					return operation.LastWeb;
				case OperationLevels.List:
					return operation.LastList.ParentWeb;
				default:
					return null; //todo throw exception
			}
		}
	}
}