using System.Collections.Generic;
using FluentlySharepoint;
using FluentlySharepoint.Extensions;
using WebProvisioning.Models;
using WebModel = WebProvisioning.Models.Web;
using ListModel = WebProvisioning.Models.List;

namespace WebProvisioning
{
	public static class Provisioning
	{
		public static CSOMOperation CreateCustomList(this CSOMOperation operation, ListModel listModel)
		{
			return operation.CreateList(listModel.Title, listModel.Template).CreateCustomFields(listModel.Columns);
		}

		public static CSOMOperation CreateCustomLists(this CSOMOperation operation, List<ListModel> listModels)
		{
			listModels.ForEach(lm => operation.CreateCustomList(lm));

			return operation;
		}

		public static CSOMOperation CreateCustomField(this CSOMOperation operation, ListColumn columnModel)
		{
			return operation.AddColumn(
				columnModel.Name,
				columnModel.Type,
				columnModel.DisplayName,
				columnModel.Required,
				columnModel.UniqueValues);
		}

		public static CSOMOperation CreateCustomFields(this CSOMOperation operation, List<ListColumn> columnModels)
		{
			columnModels.ForEach(cm => operation.CreateCustomField(cm));

			return operation;
		}

		public static CSOMOperation CreateCustomWeb(this CSOMOperation operation, WebModel webModel)
		{
			operation.CreateWeb(webModel.Name, webModel.Url, webModel.Template).CreateCustomLists(webModel.Lists);
			return operation;
		}

		public static CSOMOperation CreateCustomWebs(this CSOMOperation operation, List<WebModel> webModels)
		{
			webModels.ForEach(wm => operation.CreateCustomWeb(wm));

			return operation;
		}
	}
}