using System;
using System.Collections.Generic;
using System.Linq;
using KeenMate.FluentlySharePoint.Assets;
using KeenMate.FluentlySharePoint.Interfaces;
using KeenMate.FluentlySharePoint.Loggers;
using Microsoft.SharePoint.Client;

namespace KeenMate.FluentlySharePoint
{
	public class CSOMOperation : IDisposable
	{
		public string OriginalWebUrl { get; }
		public Web RootWeb { get; set; }
		public OperationLevels OperationLevel { get; protected set; } = OperationLevels.Web;

		private LevelLock LevelLock { get; } = new LevelLock();
		public int DefaultTimeout { get; private set; }

		public ClientContext Context { get; set; }

		public ILogger Logger { get; set; } = new BlackHoleLogger();
		public Dictionary<string, Site> LoadedSites { get; } = new Dictionary<string, Site>(5);
		public Dictionary<string, Web> LoadedWebs { get; } = new Dictionary<string, Web>(5);
		public Dictionary<string, List> LoadedLists { get; } = new Dictionary<string, List>(5);

		public Site LastSite { get; private set; }
		public Web LastWeb { get; private set; }
		public List LastList { get; private set; }

		public CSOMOperation(ClientContext context) : this(context, null)
		{
		}

		public CSOMOperation(ClientContext context, ILogger logger = null)
		{
			Context = context;
			Logger = logger??Logger;

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
		}

		private void setupOperation(ClientContext context)
		{
			DefaultTimeout = Context.RequestTimeout;

			LastSite = Context.Site;
			LastWeb = RootWeb = Context.Web;

			LoadWebRequired(LastWeb);
			LoadSiteRequired(LastSite);

			//Context.Load(LastSite);
			ActionQueue.Enqueue(new DeferredAction { ClientObject = LastSite, Action = DeferredActions.Load });

			//Context.Load(LastWeb);
			ActionQueue.Enqueue(new DeferredAction { ClientObject = LastWeb, Action = DeferredActions.Load });
		}

		public Action<ClientContext> Executor { get; set; }

		public Func<CSOMOperation, Exception, CSOMOperation> FailHandler { get; set; }

		public Queue<DeferredAction> ActionQueue { get; } = new Queue<DeferredAction>(10);

		public CSOMOperation LockLevels(params OperationLevels[] levels)
		{
			LevelLock.SetLocks(levels, true);

			return this;
		}

		public CSOMOperation UnlockLevels(params OperationLevels[] levels)
		{
			LevelLock.SetLocks(levels, false);

			return this;
		}

		public void SetLevel(OperationLevels level, ClientObject levelObject)
		{
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
			}
		}

		public void LoadSiteRequired(Site site)
		{
			Context.Load(site, s => s.ServerRelativeUrl);
		}

		public void LoadWebRequired(Web web)
		{
			Context.Load(web, w => w.ServerRelativeUrl, w => w.ListTemplates);
		}

		public void LoadListRequired(List list)
		{
			Context.Load(list, l => l.Title, l => l.RootFolder);
		}

		private void ProcessDelete(ClientObject clientObject)
		{
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
			switch (clientObject)
			{
				case Web w:
					LoadedWebs[w.ServerRelativeUrl] = w;
					break;
				case Site s:
					LoadedSites[s.ServerRelativeUrl] = s; 
					break;
				case List l:
					LoadedLists[l.Title] = l;
					break;
				case WebCollection wc:
					wc.ToList().ForEach(ProcessLoaded);
					break;
				case ListCollection lc:
					lc.ToList().ForEach(ProcessLoaded);
					break;
			}
		}

		public CSOMOperation Execute()
		{
			executeContext(out var success); // no sense to continue processing when the first execute failed

			if (!success) return this;

			while (ActionQueue.Count > 0)
			{
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
			}

			return executeContext(out success);
		}

		private CSOMOperation executeContext(out bool successful)
		{
			Logger.Debug(Messages.AboutToExecute);

			if (Executor != null)
			{
				Logger.Debug(Messages.AboutToCallExecutor);
				Executor.Invoke(Context);
			}

			try
			{
				Context.ExecuteQuery();
				Logger.Debug(Messages.SuccededToExecute);
				successful = true;
				return this;
			}
			catch (Exception ex)
			{
				Logger.Warn(ex, Messages.FailedToExecute);
				FailHandler?.Invoke(this, ex);
				successful = false;
				return this;
			}
		}

		public void Dispose()
		{
			Context?.Dispose();
		}
	}
}
