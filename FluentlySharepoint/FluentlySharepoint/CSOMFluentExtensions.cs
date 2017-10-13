using System;
using System.Linq;
using System.Net;
using System.Security;
using KeenMate.FluentlySharePoint.Interfaces;
using KeenMate.FluentlySharePoint.Models;
using Microsoft.SharePoint.Client;

namespace KeenMate.FluentlySharePoint
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

		public static CSOMOperation Create(this ClientContext context)
		{
			return new CSOMOperation(context);
		}

		public static CSOMOperation Create(this ClientContext context, ILogger logger)
		{
			return new CSOMOperation(context, logger);
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

		public static CSOMOperation SetTimeout(this CSOMOperation operation, int timeout)
		{
			operation.Context.RequestTimeout = timeout;
			return operation;
		}

		public static CSOMOperation ResetTimeout(this CSOMOperation operation)
		{
			operation.Context.RequestTimeout = operation.DefaultTimeout;
			return operation;
		}

		public static CSOMOperation OnEachRequest(this CSOMOperation operation, Action<ClientContext> executor)
		{
			operation.Executor = executor;
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
