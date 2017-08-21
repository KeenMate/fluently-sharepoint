using System;
using System.Collections.Generic;
using FluentlySharepoint.Interfaces;
using Microsoft.SharePoint.Client;

namespace FluentlySharepoint
{
	public class CSOMOperation : IDisposable
	{
		public string OriginalWebUrl { get; }
		public Web RootWeb { get; set; }
		public OperationLevels OperationLevel { get; protected set; } = OperationLevels.Web;

		private LevelLock LevelLock { get; } = new LevelLock();

		public CSOMOperation(string webUrl, ILogger logger = null)
		{
			OriginalWebUrl = webUrl;
			Context = new ClientContext(webUrl);

			Logger = logger;

			LastSite = Context.Site;
			LastWeb = RootWeb = Context.Web;

			Context.Load(LastSite);
			ActionQueue.Enqueue(new DeferredAction { ClientObject = LastSite, Action = DeferredActions.Load });

			Context.Load(LastWeb);
			ActionQueue.Enqueue(new DeferredAction { ClientObject = LastWeb, Action = DeferredActions.Load });
		}

		public ClientContext Context { get; set; }

		public ILogger Logger { get; set; }
		public Dictionary<string, Site> LoadedSites { get; } = new Dictionary<string, Site>(5);
		public Dictionary<string, Web> LoadedWebs { get; } = new Dictionary<string, Web>(5);
		public Dictionary<string, List> LoadedLists { get; } = new Dictionary<string, List>(5);

		public Site LastSite { get; private set; }
		public Web LastWeb { get; private set; }
		public List LastList { get; private set; }

		public Queue<DeferredAction> ActionQueue { get; } = new Queue<DeferredAction>(10);

		public CSOMOperation LockLevels(params OperationLevels[] levels)
		{
			SetLocks(levels, true);
			
			return this;
		}

		public CSOMOperation UnlockLevels(params OperationLevels[] levels)
		{
			SetLocks(levels, false);

			return this;
		}

		private void SetLocks(OperationLevels[] levels, bool value)
		{
			foreach (var level in levels)
			{
				switch (level)
				{
					case OperationLevels.Web:
						LevelLock.Web = value;
						break;
					case OperationLevels.Site:
						LevelLock.Site = value;
						break;
					case OperationLevels.List:
						LevelLock.List = value;
						break;
				}
			}
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

		public void Dispose()
		{
			Context?.Dispose();
		}
	}

	class LevelLock
	{
		public bool Site { get; set; }
		public bool Web { get; set; }
		public bool List { get; set; }
	}
}
