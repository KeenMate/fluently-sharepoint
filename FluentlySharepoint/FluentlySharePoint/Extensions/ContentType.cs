using System;
using System.Linq;
using System.Linq.Expressions;
using Microsoft.SharePoint.Client;

namespace KeenMate.FluentlySharePoint.Extensions
{
	public static class ContentType
	{
		//todo logging
		public static CSOMOperation LoadContentTypes(this CSOMOperation operation, params Expression<Func<Microsoft.SharePoint.Client.ContentType, object>>[] keysToLoad)
		{
			var contentTypes = DecideContentTypes(operation);

			if (contentTypes != null)
			{

				operation.Context.Load(contentTypes, ct => ct.Include(type => type.Fields, type => type.FieldLinks, type => type.Name), ct => ct.Include(keysToLoad));
			}

			return operation;
		}

		public static CSOMOperation SelectContentType(this CSOMOperation operation, string name)
		{
			var contentTypes = DecideContentTypes(operation);

			if (contentTypes != null)
			{
				operation.SetLevel(OperationLevels.ContentType, contentTypes.First(ct => ct.Name.Equals(name)));
			}

			return operation;
		}

		private static ContentTypeCollection DecideContentTypes(CSOMOperation operation)
		{
			ContentTypeCollection contentTypes = null;

			switch (operation.OperationLevel)
			{
				case OperationLevels.Web:
					contentTypes = operation.LastWeb.ContentTypes;
					break;
				case OperationLevels.Site:
					contentTypes = operation.LastSite.RootWeb.ContentTypes;
					break;
				case OperationLevels.List:
					contentTypes = operation.LastList.ContentTypes;
					break;
			}

			return contentTypes;
		}
	}
}