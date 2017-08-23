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
	}
}
