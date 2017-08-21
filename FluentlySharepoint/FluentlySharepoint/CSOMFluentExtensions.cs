using System;
using System.Linq;
using System.Net;
using System.Security;
using FluentlySharepoint.Models;
using Microsoft.SharePoint.Client;

namespace FluentlySharepoint
{
	public static class CSOMFluentExtensions
	{
		public static CSOMOperation Create(this string url)
		{
			return new CSOMOperation(url);
		}
		public static CSOMOperation SetupContext(this CSOMOperation operation, Action<ClientContext> setup)
		{
			setup(operation.Context);

			return operation;
		}

		public static CSOMOperation SetOnlineCredentials(this CSOMOperation operation, string username, SecureString password)
		{
			operation.Context.Credentials = new SharePointOnlineCredentials(username, password);

			return operation;
		}

		public static CSOMOperation LoadWebs(this CSOMOperation operation) // todo add custom loader
		{
			var webs = operation.DecideWeb().Webs;

			operation.Context.Load(webs);
			operation.ActionQueue.Enqueue(new DeferredAction { ClientObject = webs, Action = DeferredActions.Load });

			return operation;
		}

		public static CSOMOperation LoadWeb(this CSOMOperation operation, string name = "",
			Action<ClientContext, Web> webLoader = null)
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
				Url = string.IsNullOrEmpty(url) ? name : url,
				WebTemplate = template
			};

			var web = operation.DecideWeb().Webs.Add(webInformation);

			operation.Context.Load(web);
			operation.SetLevel(OperationLevels.Web, web);
			operation.ActionQueue.Enqueue(new DeferredAction { ClientObject = web, Action = DeferredActions.Load });

			return operation;
		}

		public static CSOMOperation LoadList(this CSOMOperation operation, string name, Action<ClientContext, List> listLoader = null)
		{
			var web = operation.DecideWeb();
			var list = web.Lists.GetByTitle(name);

			operation.Context.Load(web);
			if (listLoader != null)
				listLoader(operation.Context, operation.LastList);
			else
			{
				operation.Context.Load(list);
			}

			operation.SetLevel(OperationLevels.List, list);
			operation.ActionQueue.Enqueue(new DeferredAction { ClientObject = operation.LastList, Action = DeferredActions.Load });

			return operation;
		}

		public static CSOMOperation SelectList(this CSOMOperation operation, string name)
		{
			if (operation.LoadedLists.ContainsKey(name))
			{
				operation.SetLevel(OperationLevels.List, operation.LoadedLists[name]);
			}
			else
			{
				throw new ArgumentException($"List ${name} doesn't exist");
			}

			return operation;
		}

		public static CSOMOperation ChangeColumn(this CSOMOperation operation, string columnName, FieldType? type = null, string displayName = null, bool? required = null, bool? uniqueValues = null)
		{
			var field = operation.LastList.Fields.GetByInternalNameOrTitle(columnName);

			if (type.HasValue) field.TypeAsString = type.ToString();
			if (!String.IsNullOrEmpty(displayName)) field.Title = displayName;
			if (required.HasValue) field.Required = required.Value;
			if (uniqueValues.HasValue) field.EnforceUniqueValues = uniqueValues.Value;

			field.UpdateAndPushChanges(true);

			return operation;
		}

		public static CSOMOperation DeleteColumn(this CSOMOperation operation, string columnName)
		{
			var field = operation.LastList.Fields.GetByInternalNameOrTitle(columnName);
			field.DeleteObject();

			return operation;
		}

		public static CSOMOperation AddColumn(this CSOMOperation operation, string name, FieldType type, string displayName = "", bool required = false, bool uniqueValues = false)
		{
			FieldCreationInformation fieldInformation = new FieldCreationInformation
			{
				InternalName = name,
				DisplayName = String.IsNullOrEmpty(displayName) ? name : displayName,
				FieldType = type,
				Required = required,
				UniqueValues = uniqueValues
			};

			operation.LastList.Fields.AddFieldAsXml(fieldInformation.ToXml(), true, AddFieldOptions.AddFieldInternalNameHint | AddFieldOptions.AddFieldToDefaultView);

			return operation;
		}

		public static ListItemCollection GetItems(this CSOMOperation operation, string queryString)
		{
			var caml = new CamlQuery { ViewXml = queryString };

			return operation.GetItems(caml);
		}

		public static ListItemCollection GetItems(this CSOMOperation operation)
		{
			return GetItems(operation, CamlQuery.CreateAllItemsQuery());
		}

		public static ListItemCollection GetItems(this CSOMOperation operation, CamlQuery query)
		{
			var listItems = operation.LastList.GetItems(query);

			operation.Context.Load(listItems);
			operation.Execute();

			return listItems;
		}

		public static CSOMOperation DeleteItems(this CSOMOperation operation)
		{
			var caml = CamlQuery.CreateAllItemsQuery();

			operation.DeleteItems(caml);

			return operation;
		}

		public static CSOMOperation DeleteItems(this CSOMOperation operation, string queryString)
		{
			var caml = new CamlQuery { ViewXml = queryString };

			operation.DeleteItems(caml);

			return operation;
		}

		public static CSOMOperation DeleteItems(this CSOMOperation operation, CamlQuery query)
		{
			var items = operation.LastList.GetItems(query);

			operation.Context.Load(items);
			operation.ActionQueue.Enqueue(new DeferredAction { ClientObject = items, Action = DeferredActions.Delete });

			return operation;
		}

		public static CSOMOperation Execute(this CSOMOperation operation)
		{
			operation.Context.ExecuteQuery();

			foreach (var action in operation.ActionQueue)
			{
				switch (action.Action)
				{
					case DeferredActions.Load:
						operation.ProcessLoaded(action.ClientObject);
						break;
					case DeferredActions.Delete:
						operation.ProcessDelete(action.ClientObject);
						break;
				}
			}

			operation.Context.ExecuteQuery();

			return operation;
		}

		public static CSOMOperation CreateList(this CSOMOperation operation, string name, string template = null)
		{
			ListCreationInformation listInformation = new ListCreationInformation
			{
				Title = name,
				ListTemplate = String.IsNullOrEmpty(template)
					? operation.LastWeb.ListTemplates.GetByName("Custom List")
					: operation.LastWeb.ListTemplates.GetByName(template)
			};

			var list = operation.LastWeb.Lists.Add(listInformation);

			operation.LastWeb.Context.Load(list);
			operation.SetLevel(OperationLevels.List, list);
			operation.ActionQueue.Enqueue(new DeferredAction{ClientObject = list, Action = DeferredActions.Load});

			return operation;
		}

		public static CSOMOperation DeleteList(this CSOMOperation operation, string name)
		{
			var list = operation.LastWeb.Lists.GetByTitle(name);

			operation.ActionQueue.Enqueue(new DeferredAction { ClientObject = list, Action = DeferredActions.Delete });

			return operation;
		}

		public static CSOMOperation DeleteWeb(this CSOMOperation operation)
		{
			operation.LastWeb.DeleteObject();

			operation.ActionQueue.Enqueue(new DeferredAction {ClientObject = operation.LastWeb, Action = DeferredActions.Delete});
			return operation;
		}

		private static void ProcessDelete(this CSOMOperation operation, ClientObject clientObject)
		{
			switch (clientObject)
			{
				case Web w:
					operation.LoadedWebs.Remove(w.Url);
					w.DeleteObject();
					break;
				case List l:
					operation.LoadedLists.Remove(l.Title);
					l.DeleteObject();
					break;
				case ListItemCollection lic:
					lic.ToList().ForEach(li => li.DeleteObject());
					break;
			}
		}

		private static void ProcessLoaded(this CSOMOperation operation, ClientObject clientObject)
		{
			switch (clientObject)
			{
				case Web w:
					if (!operation.LoadedWebs.ContainsKey(w.ServerRelativeUrl))
						operation.LoadedWebs.Add(w.ServerRelativeUrl, w);
					break;
				case Site s:
					if (!operation.LoadedSites.ContainsKey(s.ServerRelativeUrl))
						operation.LoadedSites.Add(s.ServerRelativeUrl, s);
					break;
				case List l:
					if (!operation.LoadedLists.ContainsKey(l.Title))
						operation.LoadedLists.Add(l.Title, l);
					break;
				case WebCollection wc:
					wc.ToList().ForEach(operation.ProcessLoaded);
					break;
				case ListCollection lc:
					lc.ToList().ForEach(operation.ProcessLoaded);
					break;
			}
		}

		private static Web DecideWeb(this CSOMOperation operation)
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
