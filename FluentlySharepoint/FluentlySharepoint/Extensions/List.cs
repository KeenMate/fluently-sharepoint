using System;
using System.Linq;
using KeenMate.FluentlySharePoint.Models;
using Microsoft.SharePoint.Client;
using KeenMate.FluentlySharePoint.Assets;
using KeenMate.FluentlySharePoint.Enums;
using ListTemplate = KeenMate.FluentlySharePoint.Enums.ListTemplate;

namespace KeenMate.FluentlySharePoint.Extensions
{
	public static class List
	{
		public static CSOMOperation LoadList(this CSOMOperation operation, string name, Action<ClientContext, Microsoft.SharePoint.Client.List> listLoader = null)
		{
			operation.LogDebug($"Loading list {name}");

			var web = operation.DecideWeb();
			var list = web.Lists.GetByTitle(name);

			operation.LoadListRequired(list);

			if (listLoader != null)
				listLoader(operation.Context, list);
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

		public static CSOMOperation ModifyColumn(this CSOMOperation operation, string columnName, FieldType? type = null, string displayName = null, bool? required = null, bool? uniqueValues = null)
		{
			operation.LogInfo($"Modifying column {columnName}");

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
			operation.LogInfo($"Removing column {columnName}");

			var field = operation.LastList.Fields.GetByInternalNameOrTitle(columnName);
			field.DeleteObject();

			return operation;
		}

		public static CSOMOperation AddColumn(this CSOMOperation operation, string name, FieldType type, string displayName = "", bool required = false, bool uniqueValues = false)
		{
			operation.LogInfo($"Adding column {name}");

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

		public static ListItemCollection GetItems(this CSOMOperation operation, string queryString, int? rowLimit = null)
		{
			if (rowLimit != null)
				queryString = string.Format(CamlQueries.WrappedWithRowLimit, queryString, rowLimit);

			var caml = new CamlQuery { ViewXml = $"<View>{queryString}</View>" };

			return operation.GetItems(caml);
		}

		public static ListItemCollection GetItems(this CSOMOperation operation)
		{
			return GetItems(operation, CamlQuery.CreateAllItemsQuery());
		}

		public static ListItemCollection GetItems(this CSOMOperation operation, CamlQuery query)
		{
			operation.LogInfo("Getting items");
			operation.LogDebug("Query:\n{query.ViewXml}");

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
			operation.LogInfo("Deleting items");
			operation.LogDebug($"Query:\n{query}");

			var items = operation.LastList.GetItems(query);

			operation.Context.Load(items);
			operation.ActionQueue.Enqueue(new DeferredAction { ClientObject = items, Action = DeferredActions.Delete });

			return operation;
		}

		public static CSOMOperation CreateList(this CSOMOperation operation, string name, string template = null)
		{
			operation.LogInfo($"Creating list {name}");

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

		public static CSOMOperation CreateList(this CSOMOperation operation, string name, ListTemplate template)
		{
			return operation.CreateList(name, operation.LastWeb.ListTemplates.First(t => t.ListTemplateTypeKind == (int) template).Name);
		}

		public static CSOMOperation DeleteList(this CSOMOperation operation, string name)
		{
			operation.LogInfo($"Deleting list {name}");

			var list = operation.LoadedLists[name];

			operation.ActionQueue.Enqueue(new DeferredAction { ClientObject = list, Action = DeferredActions.Delete });

			return operation;
		}
	}
}