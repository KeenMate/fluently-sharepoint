using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using KeenMate.FluentlySharePoint.Assets;
using KeenMate.FluentlySharePoint.Enums;
using KeenMate.FluentlySharePoint.Helpers;
using KeenMate.FluentlySharePoint.Interfaces;
using KeenMate.FluentlySharePoint.Loggers;
using KeenMate.FluentlySharePoint.Models;
using KeenMate.FluentlySharePoint.Models.Taxonomy;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using ListTemplate = Microsoft.SharePoint.Client.ListTemplate;

namespace KeenMate.FluentlySharePoint
{
	public class CSOMOperation : IDisposable
	{
		public static class DefaultRetrievals
		{
			public static Expression<Func<Site, object>>[] Site = new Expression<Func<Site, object>>[]
			{
				g => g.Id,
				g => g.ServerRelativeUrl,
				g => g.ReadOnly,
				g => g.Url
			};

			public static Expression<Func<WebCollection, object>>[] WebCollection =
				new Expression<Func<WebCollection, object>>[]
				{
					g => g.Include(w => Web)
				};

			public static Expression<Func<Web, object>>[] Web = new Expression<Func<Web, object>>[]
			{
				g => g.Id,
				g => g.ServerRelativeUrl,
				g => g.Title,
				g => g.Url,
				g => g.Description,
				g => g.Language,
				g => g.IsMultilingual,
				g => g.WebTemplate
			};

			public static Expression<Func<List, object>>[] List = new Expression<Func<List, object>>[]
			{
				g => g.Id,
				g => g.Title,
				g => g.Description,
				g => g.EnableAssignToEmail,
				g => g.EnableAttachments,
				g => g.EnableFolderCreation,
				g => g.EnableMinorVersions,
				g => g.EnableModeration,
				g => g.EnableVersioning,
				g => g.Hidden,
				g => g.IsApplicationList,
				g => g.IsCatalog,
				g => g.IsEnterpriseGalleryLibrary,
				g => g.IsPrivate,
				g => g.IsSiteAssetsLibrary,
				g => g.IsSystemList,
				g => g.ItemCount,
				g => g.RootFolder
			};

			public static Expression<Func<WebTemplateCollection, object>>[] WebTemplateCollection =
				new Expression<Func<WebTemplateCollection, object>>[]
				{
					g => g.Include(WebTemplate)
				};

			public static Expression<Func<WebTemplate, object>>[] WebTemplate = new Expression<Func<WebTemplate, object>>[]
			{
				g => g.Title,
				g => g.Name,
				g => g.Description,
				g => g.DisplayCategory,
				g => g.Id,
				g => g.IsRootWebOnly,
				g => g.IsSubWebOnly,
				g => g.Lcid
			};

			public static Expression<Func<ListTemplateCollection, object>>[] ListTemplateCollection =
				new Expression<Func<ListTemplateCollection, object>>[]
				{
					g => g.Include(t => ListTemplate)
				};

			public static Expression<Func<ListTemplate, object>>[] ListTemplate = new Expression<Func<ListTemplate, object>>[]
			{
				g => g.Name,
				g => g.InternalName,
				g => g.Description,
				g => g.AllowsFolderCreation,
				g => g.IsCustomTemplate,
				g => g.BaseType
			};

		}

		public static string[] UrlCharsToRemove =
			new[] { "\"", "\"", "`", " ", "?", "!", "@", "#", "$", "%", "^", "*", "(", ")", "[", "]", "<", ">" };

		public static string[] UrlCharsToReplace = new[] { " ", ".", ",", "/", "\\", ";", "&", "|", };
		public static string UrlReplaceChar = "-";

		public static Func<string, string> UrlNormalizeFunctor = (title) =>
		{
			title = title.RemoveDiacritics();

			foreach (var c in UrlCharsToRemove)
			{
				title = title.Replace(c, UrlReplaceChar);
			}

			foreach (var c in UrlCharsToReplace)
			{
				title = title.Replace(c, UrlReplaceChar);

			}
			// not proud of this one, please, forgive me :)
			return title;
		};

		public uint DefaultLcid { get; private set; } = (uint)Lcid.English;
		public int DefaultCompatibilityLevel { get; private set; } = (int)SharePointVersions.SP2016;

		public string OriginalWebUrl { get; }
		public Web RootWeb { get; set; }
		public OperationLevels OperationLevel { get; protected set; } = OperationLevels.Web;

		private LevelLock LevelLock { get; } = new LevelLock();
		public int DefaultTimeout { get; private set; }
		public bool ThrowOnError { get; private set; }

		public ClientContext Context { get; set; }

		public Guid CorrelationId { get; set; }
		private ILogger Logger { get; set; } = new BlackHoleLogger();
		public Func<Guid, string, string> LogMessageFormat { get; set; } =
			(correlationId, message) => $"{(correlationId != Guid.Empty ? $"{correlationId}: " : "")}{message}";

		public Dictionary<string, Site> LoadedSites { get; } = new Dictionary<string, Site>(5);
		public Dictionary<string, Web> LoadedWebs { get; } = new Dictionary<string, Web>(5);
		public Dictionary<string, List> LoadedLists { get; } = new Dictionary<string, List>(5);

		public Site LastSite { get; private set; }
		public Web LastWeb { get; private set; }
		public List LastList { get; private set; }
		public ContentType LastContentType { get; private set; }

		public TaxonomyOperation TaxonomyOperation { get; set; }

		public CSOMOperation(ClientContext context) : this(context, null)
		{}

		public CSOMOperation(ClientContext context, ILogger logger = null)
		{
			Context = context;
			Logger = logger ?? Logger;

			setupOperation(Context);
		}

		public CSOMOperation(string webUrl)
		{
			OriginalWebUrl = webUrl;
			Context = new ClientContext(webUrl);

			setupOperation(Context);
		}

		public CSOMOperation(string webUrl, ILogger logger = null) : this(webUrl)
		{
			Logger = logger ?? Logger;
			LogInfo("CSOM Operation created and ready to rock'n'roll");
		}

		private void setupOperation(ClientContext context)
		{
			LogDebug($"Setting default timeout to: {Context.RequestTimeout}");
			DefaultTimeout = Context.RequestTimeout;

			LastSite = Context.Site;
			LastWeb = RootWeb = Context.Web;

			LogDebug("Loading initial data");

			LoadWebWithDefaultRetrievals(LastWeb);
			LoadSiteWithDefaultRetrievals(LastSite);

			ActionQueue.Enqueue(new DeferredAction { ClientObject = LastSite, Action = DeferredActions.Load });
			ActionQueue.Enqueue(new DeferredAction { ClientObject = LastWeb, Action = DeferredActions.Load });
		}

		/// <summary>
		/// Global on being executed handler
		/// </summary>
		public Action<ClientContext> OnBeingExecuted { get; set; }

		/// <summary>
		/// Global on fail handler
		/// </summary>
		public Func<CSOMOperation, Exception, CSOMOperation> OnFail { get; private set; }

		/// <summary>
		/// Should an exception be rethrown on execution failure
		/// </summary>
		/// <param name="yesNo"></param>
		/// <returns></returns>
		public CSOMOperation ThrowExceptionOnError(bool yesNo)
		{
			ThrowOnError = yesNo;
			return this;
		}

		#region Logging

		public void LogTrace(string message) => Logger.Trace(LogMessageFormat(CorrelationId, message));
		public void LogDebug(string message) => Logger.Debug(LogMessageFormat(CorrelationId, message));
		public void LogInfo(string message) => Logger.Info(LogMessageFormat(CorrelationId, message));
		public void LogWarn(string message) => Logger.Warn(LogMessageFormat(CorrelationId, message));
		public void LogError(string message) => Logger.Error(LogMessageFormat(CorrelationId, message));
		public void LogFatal(string message) => Logger.Fatal(LogMessageFormat(CorrelationId, message));


		#endregion

		public CSOMOperation SetDefaultLCID(uint lcid)
		{
			DefaultLcid = lcid;
			return this;
		}

		public CSOMOperation SetDefaultCompatibilityLevel(int compatibilityLevel)
		{
			DefaultCompatibilityLevel = compatibilityLevel;
			return this;
		}

		public Queue<DeferredAction> ActionQueue { get; } = new Queue<DeferredAction>(10);

		public CSOMOperation LockLevels(params OperationLevels[] levels)
		{
			LogTrace($"Locking levels to: {levels}");

			LevelLock.SetLocks(levels, true);

			return this;
		}

		public CSOMOperation UnlockLevels(params OperationLevels[] levels)
		{
			LogTrace($"Unlocking levels: {levels}");

			LevelLock.SetLocks(levels, false);

			return this;
		}

		public void SetLevel(OperationLevels level, ClientObject levelObject)
		{
			LogTrace($"Setting operation level to: {level}");
			switch (levelObject)
			{
				case Site s when level == OperationLevels.Site && !LevelLock.Site:
					LastSite = s;
					OperationLevel = level;
					break;
				case Web w when level == OperationLevels.Web && !LevelLock.Web:
					LastWeb = w;
					OperationLevel = level;
					break;
				case List l when level == OperationLevels.List && !LevelLock.List:
					LastList = l;
					OperationLevel = OperationLevels.List;
					break;
				case ContentType c when level == OperationLevels.ContentType && !LevelLock.ContentType:
					LastContentType = c;
					OperationLevel = OperationLevels.ContentType;
					break;
			}
		}

		public void LoadSiteWithDefaultRetrievals(Site site)
		{
			LogTrace($"Loading site with default retrievals");

			Context.Load(site, s => s.ServerRelativeUrl);
		}

		public void LoadWebWithDefaultRetrievals(Web web)
		{
			LogTrace($"Loading web with default retrievals");
			Context.Load(web, DefaultRetrievals.Web);
		}

		public void LoadListRequired(List list)
		{
			LogTrace($"Loading list with default retrievals");

			Context.Load(list, DefaultRetrievals.List);
		}

		private void ProcessDelete(ClientObject clientObject)
		{
			LogTrace("Processing deleted object");
			switch (clientObject)
			{
				case Web w:
					LoadedWebs.Remove(w.Url);
					w.DeleteObject();
					break;
				case List l:
					LoadedLists.Remove(l.Title);
					l.DeleteObject();
					break;
				case ListItemCollection lic:
					lic.ToList().ForEach(li => li.DeleteObject());
					break;
			}
		}

		private void ProcessLoaded(ClientObject clientObject)
		{
			LogTrace("Processing loaded object");
			switch (clientObject)
			{
				case Web w:
					LoadedWebs.AddOrUpdate(w.ServerRelativeUrl, w);
					break;
				case Site s:
					LoadedSites.AddOrUpdate(s.ServerRelativeUrl, s);
					break;
				case List l:
					LoadedLists.AddOrUpdate(l.Title, l);
					break;
				case WebCollection wc:
					wc.ToList().ForEach(ProcessLoaded);
					break;
				case ListCollection lc:
					lc.ToList().ForEach(ProcessLoaded);
					break;
			}
		}

		public CSOMOperation SetLogMessageFormat(Func<Guid, string, string> logMessageFormat)
		{
			LogMessageFormat = logMessageFormat;
			return this;
		}

		public CSOMOperation Load<T>(T clientObject, params Expression<Func<T, object>>[] retrievals) where T : ClientObject
		{
			LogTrace($"Loading object of type: {clientObject.GetType().Name}");
			Context.Load(clientObject, retrievals);

			return this;
		}

		public CSOMOperation Execute(Func<Exception, CSOMOperation> localFailHandler = null)
		{
			LogInfo(Messages.AboutToExecute);

			executeContext(localFailHandler, out var success); // no sense to continue processing when the first execute failed

			if (!success) return this;

			LogDebug($"Items in action queue: {ActionQueue.Count}");
			var i = 1;
			while (ActionQueue.Count > 0)
			{
				LogDebug($"Processing action: {i} of {ActionQueue.Count}");
				var action = ActionQueue.Dequeue();

				switch (action.Action)
				{
					case DeferredActions.Load:
						ProcessLoaded(action.ClientObject);
						break;
					case DeferredActions.Delete:
						ProcessDelete(action.ClientObject);
						break;
				}

				i++;
			}

			return executeContext(localFailHandler, out success);
		}

		public string NormalizeUrl(string title)
		{
			return UrlNormalizeFunctor.Invoke(title);
		}

		private CSOMOperation executeContext(Func<Exception, CSOMOperation> localFailHandler, out bool successful)
		{
			LogTrace($"ThrowOnError set to: {ThrowOnError}, OnBeingExecuted defined: {OnBeingExecuted != null}, LocalFailHandler defined: {localFailHandler != null}, OnFail defined: {OnFail != null}");
			if (OnBeingExecuted != null)
			{
				LogDebug(Messages.AboutToCallExecutor);
				OnBeingExecuted.Invoke(Context);
			}

			try
			{
				LogTrace("Executing context");
				Context.ExecuteQuery();
				LogDebug(Messages.SuccededToExecute);
				successful = true;
				return this;
			}
			catch (Exception ex)
			{
				LogWarn(string.Format(Messages.FailedToExecute, ex.Message));
				successful = false;

				if (localFailHandler != null)
				{
					LogTrace("Calling local fail handler");
					localFailHandler.Invoke(ex);
					return this;
				}

				if (ThrowOnError)
					throw;

				OnFail?.Invoke(this, ex);

				return this;
			}
		}

		public void Dispose()
		{
			Context?.Dispose();
		}
	}
}
