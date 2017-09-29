using System;
using System.Linq;
using System.Net;
using System.Security;
using FluentlySharepoint.Interfaces;
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

		public static CSOMOperation Create(this string url, ILogger logger)
		{
			return new CSOMOperation(url, logger);
		}

		public static CSOMOperation SetupContext(this CSOMOperation operation, Action<ClientContext> setup)
		{
			setup(operation.Context);

			return operation;
		}

		public static CSOMOperation SetOnlineCredentials(this CSOMOperation operation, string username, string password)
		{
			return operation.SetOnlineCredentials(username, password.ToSecureString());
		}

		public static CSOMOperation SetOnlineCredentials(this CSOMOperation operation, string username, SecureString password)
		{
			operation.Context.Credentials = new SharePointOnlineCredentials(username, password);

			return operation;
		}

		/// <summary>
		/// On fail handler executed in all-catch block of clientContext.Execute() command
		/// </summary>
		/// <param name="operation">This operation</param>
		/// <param name="handler">Handler that is assigned to CSOMOperation.FailHandler property</param>
		/// <returns>This operation</returns>
		public static CSOMOperation Fail(this CSOMOperation operation, Func<CSOMOperation, Exception, CSOMOperation> handler)
		{
			operation.FailHandler = handler;
			return operation;
		}
	}
}
