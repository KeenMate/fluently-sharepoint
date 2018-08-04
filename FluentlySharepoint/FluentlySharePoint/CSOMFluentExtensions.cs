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
		/// <summary>
		/// Create a client context directly from URL string
		/// </summary>
		/// <param name="url">Usual http/s URL address</param>
		/// <returns></returns>
		public static CSOMOperation Create(this string url)
		{
			return new CSOMOperation(url);
		}

		/// <summary>
		/// Create a client context directly from URL string with custom logger 
		/// </summary>
		/// <param name="url">Usual http/s URL address</param>
		/// <param name="logger">Custom logger</param>
		/// <returns></returns>
		public static CSOMOperation Create(this string url, ILogger logger)
		{
			return new CSOMOperation(url, logger);
		}

		/// <summary>
		/// To use with already created client context
		/// </summary>
		/// <param name="context"></param>
		/// <remarks>You can use this for example in SharePoint Add-ins where the context is created for you by the template.</remarks>
		/// <returns></returns>
		public static CSOMOperation Create(this ClientContext context)
		{
			return new CSOMOperation(context);
		}

		/// <summary>
		/// To use with already created client context
		/// </summary>
		/// <param name="context"></param>
		/// <param name="logger">ILogger instance</param>
		/// <remarks>You can use this for example in SharePoint Add-ins where the context is created for you by the template.</remarks>
		/// <returns></returns>
		public static CSOMOperation Create(this ClientContext context, ILogger logger)
		{
			return new CSOMOperation(context, logger);
		}

		/// <summary>
		/// Setup client context as you need
		/// </summary>
		/// <param name="operation"></param>
		/// <param name="setup"></param>
		/// <returns></returns>
		public static CSOMOperation SetupContext(this CSOMOperation operation, Action<ClientContext> setup)
		{
			setup(operation.Context);

			return operation;
		}

		/// <summary>
		/// Sets authentication mode for the context
		/// </summary>
		/// <param name="operation"></param>
		/// <param name="mode">Desired <see cref="ClientAuthenticationMode"/></param>
		/// <returns></returns>
		public static CSOMOperation SetAuthenticationMode(this CSOMOperation operation, ClientAuthenticationMode mode)
		{
			operation.Context.AuthenticationMode = mode;

			return operation;
		}

	/// <summary>
		/// Set online credentials with username and plain password
		/// </summary>
		/// <param name="operation"></param>
		/// <param name="username"></param>
		/// <param name="password">Plain string password, automatically converted to SecureString</param>
		/// <returns></returns>
		public static CSOMOperation SetOnlineCredentials(this CSOMOperation operation, string username, string password)
		{
			return operation.SetOnlineCredentials(username, password.ToSecureString());
		}

		/// <summary>
		/// Set online credentials with username and already secured password
		/// </summary>
		/// <param name="operation"></param>
		/// <param name="username"></param>
		/// <param name="password">Secured string password</param>
		/// <returns></returns>
		public static CSOMOperation SetOnlineCredentials(this CSOMOperation operation, string username, SecureString password)
		{
			operation.SetAuthenticationMode(ClientAuthenticationMode.Default);
			operation.LogDebug("Setting SharePoint Online credentials");
			operation.Context.Credentials = new SharePointOnlineCredentials(username, password);

			return operation;
		}

		public static CSOMOperation SetNetworkCredentials(this CSOMOperation operation, string username, string password)
		{
			operation.Context.Credentials = new NetworkCredential(username, password);
			return operation;
		}

		public static CSOMOperation SetNetworkCredentials(this CSOMOperation operation, string domain, string username,
			string password)
		{
			operation.Context.Credentials = new NetworkCredential(username, password, domain);
			return operation;
		}

		/// <summary>
		/// Set client context operation timeout
		/// </summary>
		/// <param name="operation"></param>
		/// <param name="timeout">In miliseconds</param>
		/// <returns></returns>
		public static CSOMOperation SetTimeout(this CSOMOperation operation, int timeout)
		{
			operation.LogDebug($"Setting request timeout to {timeout}");
			operation.Context.RequestTimeout = timeout;
			return operation;
		}

		/// <summary>
		/// Reset client context operation timeout to default
		/// </summary>
		/// <param name="operation"></param>
		/// <returns></returns>
		public static CSOMOperation ResetTimeout(this CSOMOperation operation)
		{
			operation.LogDebug("Setting request timeout to default");
			operation.Context.RequestTimeout = operation.DefaultTimeout;
			return operation;
		}

		/// <summary>
		/// Define a method that is called on each request
		/// </summary>
		/// <param name="operation"></param>
		/// <param name="executor"></param>
		/// <returns></returns>
		public static CSOMOperation OnEachRequest(this CSOMOperation operation, Action<ClientContext> executor)
		{
			operation.LogDebug("Operation executor set");
			operation.OnBeingExecuted = executor;
			return operation;
		}

		/// <summary>
		/// On fail handler executed in all-catch block of clientContext.Execute() command
		/// </summary>
		/// <param name="operation">This operation</param>
		/// <param name="handler">Handler that is assigned to CSOMOperation.OnFail property</param>
		/// <returns>This operation</returns>
		public static CSOMOperation Fail(this CSOMOperation operation, Func<CSOMOperation, Exception, CSOMOperation> handler)
		{
			operation.LogDebug("Fail handler set");
			operation.OnFail = handler;
			return operation;
		}
	}
}
