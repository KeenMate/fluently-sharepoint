using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using FluentlySharepoint;
using FluentlySharepoint.Extensions;
using FluentlySharePoint_Nlog;

namespace TestConsole
{
	class Program
	{
		public const string UserName = "trent.reznor@keenmate.com";
		public const string Password = "";
		public const string SiteUrl =
				"https://keenmate.sharepoint.com/sites/demo/fluently-sharepoint/";

		public static NlogLogger logger = new NlogLogger();

		static void Main(string[] args)
		{
			CreateAndExecute();

			CreateExecuteAndReuse();
		}

		private static void CreateAndExecute()
		{
			logger.CorrelationId = Guid.NewGuid();
			var op = SiteUrl
				.Create(logger)
				.SetOnlineCredentials(UserName, Password) // Available also with SecureString parameter
				.Execute();
		}

		private static void CreateExecuteAndReuse()
		{
			logger.CorrelationId = Guid.NewGuid();

			logger.Info("Create, execute and reuse example");

			var op = SiteUrl
				.Create(logger)
				.SetOnlineCredentials(UserName, Password) // Available also with SecureString parameter
				.Execute();

			var listTitle = "Documents";

			op.LoadList(listTitle, (context, list) =>
			{
				context.Load(list, l=>l.ItemCount);
			});
			
			var items = op.GetItems();

			logger.Info($"Total items of list {listTitle} with list.ItemCount: {op.LastList.ItemCount} = Items count loaded with GetItems: {items.Count}");
		}
	}
}
