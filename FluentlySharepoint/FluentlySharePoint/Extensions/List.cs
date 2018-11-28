using System;
using System.Linq;
using System.Linq.Expressions;
using Microsoft.SharePoint.Client;
using KeenMate.FluentlySharePoint.Assets;

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


		/// <summary>
		/// Get items from the last loaded list using CAML query
		/// </summary>
		/// <param name="operation"></param>
		/// <param name="queryString">CAML query string used for selection
		/// When <paramref name="rowLimit"/> is used the <paramref name="queryString"/> is wrapped into View tag with Scope="RecursiveAll".
		/// Also when there is not View tag in the <paramref name="queryString"/> it is wrapped into simple View tag with Scope="RecursiveAll" 
		/// </param>
		/// <param name="rowLimit"></param>
		/// <returns>Loaded list items in standard CSOM <see cref="ListItemCollection"/></returns>
		/// <remarks>
		/// If you need bigger control over the CAML query use alternative <seealso cref="GetItems(KeenMate.FluentlySharePoint.CSOMOperation,CamlQuery)"/> method with CamlQuery parameter
		/// </remarks>
		public static ListItemCollection GetItems(this CSOMOperation operation, string queryString, int? rowLimit = null, params Expression<Func<ListItem, object>>[] retrievals)
		{
			string caml = queryString;
			if (rowLimit != null)
				caml = string.Format(CamlQueries.WrappedWithRowLimit, queryString, rowLimit);
			else
			{
				if (!caml.ToLower().Contains("<view>"))
					caml = $"<View Scope=\"RecursiveAll\">{queryString}</View>";
			}
			var ca = new CamlQuery { ViewXml = caml };

			return operation.GetItems(ca, retrievals);
		}

		/// <summary>
		/// Get all items from the last loaded list using standard <see cref="CamlQuery.CreateAllItemsQuery()"/>
		/// </summary>
		/// <param name="operation"></param>
		/// <returns>Loaded list items in standard CSOM <see cref="ListItemCollection"/></returns>
		public static ListItemCollection GetItems(this CSOMOperation operation, params Expression<Func<ListItem, object>>[] retrievals)
		{
			return GetItems(operation, CamlQuery.CreateAllItemsQuery(), retrievals);
		}

		/// <summary>
		/// Get items from the last loaded list using standard <see cref="CamlQuery"/>
		/// </summary>
		/// <param name="operation">Beware! Context executing method</param>
		/// <param name="query">Query used in GetItems method</param>
		/// <param name="retrievals"></param>
		/// <returns>Loaded list items in standard CSOM <see cref="ListItemCollection"/></returns>
		public static ListItemCollection GetItems(this CSOMOperation operation, CamlQuery query, params Expression<Func<ListItem, object>>[] retrievals)
		{
			operation.LogInfo("Getting items");
			operation.LogDebug($"Query:\n{query.ViewXml}");

			var listItems = operation.LastList.GetItems(query);

			if (retrievals.Length > 0)
				operation.Context.Load(listItems, collection => collection.Include(retrievals));
			else
				operation.Context.Load(listItems);

			operation.Execute();

			return listItems;
		}

		/// <summary>
		/// Get Item by id
		/// </summary>
		/// <param name="operation"></param>
		/// <param name="itemId"></param>
		/// <param name="retrievals"></param>
		/// <returns></returns>
		public static ListItem GetItem(this CSOMOperation operation, int itemId, params Expression<Func<ListItem, object>>[] retrievals)
		{
			ListItem item = operation.LastList.GetItemById(itemId);

			operation
				.Load(item, retrievals.Length > 0
										? retrievals
										: CSOMOperation.DefaultRetrievals.ListItem)
				.Execute();

			return item;
		}

		/// <summary>
		/// Get Item by its title
		/// </summary>
		/// <param name="operation"></param>
		/// <param name="title"></param>
		/// <param name="retrievals"></param>
		/// <returns></returns>
		public static ListItem GetItem(this CSOMOperation operation, string title, params Expression<Func<ListItem, object>>[] retrievals)
		{
			throw new NotImplementedException();

			//CamlQuery titleQuery = new CamlQuery
			//{
			//	ViewXml = "<View><Query><Where>" +
			//							"<Eq>" +
			//								"<FieldRef Name='Title'/>" +
			//								$"<Value Type='Text'>{title}</Value>" +
			//							"</Eq>" +
			//						"</Where></Query></View>"
			//};

			//return operation.GetItem(titleQuery, retrievals);
		}

		/// <summary>
		/// Get Item using CamlQuery
		/// </summary>
		/// <param name="operation"></param>
		/// <param name="query"></param>
		/// <param name="retrievals"></param>
		/// <returns></returns>
		public static ListItem GetItem(this CSOMOperation operation, CamlQuery query)
		{
			ListItemCollection items = operation
				.LastList
				.GetItems(query);

			ListItem item = items.FirstOrDefault();

			operation.Load(item);

			return item;
		}

		public static CSOMOperation DeleteItem(this CSOMOperation operation, int itemId)
		{
			operation.LastList
				.GetItemById(itemId)
				.DeleteObject();

			return operation;
		}

		/// <summary>
		/// Remove all items from list
		/// </summary>
		/// <param name="operation"></param>
		/// <returns></returns>
		public static CSOMOperation DeleteItems(this CSOMOperation operation)
		{
			var caml = CamlQuery.CreateAllItemsQuery();

			operation.DeleteItems(caml);

			return operation;
		}

		public static CSOMOperation DeleteItems(this CSOMOperation operation, string queryString)
		{
			var caml = new CamlQuery { ViewXml = $"<View>{queryString}</View>" };

			operation.DeleteItems(caml);

			return operation;
		}

		/// <summary>
		/// This method first loads all items valid for <paramref name="query"/> and then enqueue removal actions
		/// </summary>
		/// <param name="operation">Not context execution method</param>
		/// <param name="query"><see cref="CamlQuery"/> parameter used for list item selection</param>
		/// <returns></returns>
		public static CSOMOperation DeleteItems(this CSOMOperation operation, CamlQuery query)
		{
			operation.LogInfo("Deleting items");
			operation.LogDebug($"Query:\n{query}");

			var items = operation.LastList.GetItems(query);
			operation.Context.Load(items);
			operation.ActionQueue.Enqueue(new DeferredAction { ClientObject = items, Action = DeferredActions.Delete });

			return operation;
		}
	}
}