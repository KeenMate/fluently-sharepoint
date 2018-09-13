using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Schema;
using System.Xml.Serialization;
using KeenMate.FluentlySharePoint;
using KeenMate.FluentlySharePoint.Enums;
using KeenMate.FluentlySharePoint.Extensions;
using KeenMate.FluentlySharePoint.Interfaces;
using KeenMate.FluentlySharePoint.Loggers;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WorkflowServices;
using TestConsole.WorkflowModels;

namespace TestConsole
{
	class Program
	{
		public const string SiteUrl =
				"https://keenmate.sharepoint.com/sites/demo/fluently-sharepoint/";
		
		public static ILogger logger = new ConsoleLogger();

		static void Main(string[] args)
		{
			//MeasuredOperation(CreateAndExecute);

			//MeasuredOperation(CreateExecuteAndReuse);

			//MeasuredOperation(ReuseExistingContext);

			//MeasuredOperation(CreateWebAndSeveralListInIt);

			//MeasuredOperation(StartStandardWorkflow);

			//MeasuredOperation(CreateTermSetAndTerm);
		}

		private static void MeasuredOperation(Action operation)
		{
			var stopwatch = Stopwatch.StartNew();

			operation();

			stopwatch.Stop();

			logger.Trace($"Operation finished in {stopwatch.ElapsedMilliseconds}");
		}

		private static void CreateAndExecute()
		{
			var op = SiteUrl
				.Create(logger)
				.SetOnlineCredentials(ClientSecrets.Username, ClientSecrets.Password) // Available also with SecureString parameter
				.Execute();

			logger.Trace($"Default timeout: {op.Context.RequestTimeout}");
		}

		private static void CreateExecuteAndReuse()
		{
			logger.Info("Create, execute and reuse example");

			var op = SiteUrl
				.Create(logger)
				.SetOnlineCredentials(ClientSecrets.Username, ClientSecrets.Password) // Available also with SecureString parameter
				.Execute();

			var listTitle = "Documents";

			op.LoadList(listTitle, (context, list) =>
			{
				context.Load(list, l => l.ItemCount);
			});

			var items = op.GetItems();

			logger.Info($"Total items of list {listTitle} with list.ItemCount: {op.LastList.ItemCount} = Items count loaded with GetItems: {items.Count}");
		}

		private static void ReuseExistingContext()
		{
			logger.Info("Reuse existing context example");

			ClientContext context = new ClientContext(SiteUrl);
			context.Credentials = new SharePointOnlineCredentials(ClientSecrets.Username, ClientSecrets.Password.ToSecureString());

			var listTitle = "Documents";

			var items = context
				.Create()
				.LoadList(listTitle)
				.GetItems();

			logger.Info($"Total items of list {listTitle} with list.ItemCount: {items.Count}");

		}

		private static void CreateWebAndSeveralListInIt()
		{
			var op = SiteUrl
				.Create(new ConsoleLogger())
				.SetOnlineCredentials(ClientSecrets.Username, ClientSecrets.Password)
				.CreateWeb($"New Web - {DateTime.Now:HH-mm}", (int)Lcid.English, $"NewWeb-{DateTime.Now:HH-mm}")
				.CreateList("Customers")
				.AddNumberField("Internal number")
				.AddBooleanField("EU Company")
				.AddTextField("Tax ID")
				.Execute()
				.CreateList("Customer contact")
				.AddTextField("Last name")
				.AddTextField("First name")
				.AddTextField("Email", required: true)
				.AddChoiceField("Contact level", new List<string>() { "Owner", "His wife!", "Poor employee" }, ChoiceTypes.Dropdown)
				.AddLookupField("Company", "Customers", "Title", "Related to company", required: true)
				.Execute();
		}

		/// <summary>
		/// Not working. We cannot find out a correct way how to send recepients to the standard workflow for it to work
		/// </summary>
		private static void StartStandardWorkflow()
		{
			var op = "http://dev-sp2016-01:7100/"
				.Create(new ConsoleLogger())
				.SetupContext(context =>
				{
					context.Credentials =
						new NetworkCredential() { Domain = ClientSecrets.Domain, Password = ClientSecrets.Password, UserName = ClientSecrets.Username };
				});
			//.SetOnlineCredentials(ClientSecrets.Username, ClientSecrets.Password);
			op.Fail((operation, exception) =>
			{
				Console.WriteLine(exception.Message);
				return operation;
			}).LoadList("Documents with Workflow", (context, list) =>
			{
				context.Load(list, l=>l.WorkflowAssociations, l=>l.Id);
			}).Execute();

			var x = op.LastList.WorkflowAssociations;
			op.Context.Load(x);
			var items = op.LastList.GetItems(new CamlQuery());
			
			op.Context.Load(items);
			op.Execute();

			var itemGuid = items[items.Count-1]["GUID"].ToString();

			var itemId = new Guid(itemGuid);

			var manager = new WorkflowServicesManager(op.Context, op.LastWeb);
			var instanceService = manager.GetWorkflowInstanceService();
			var interOpService = manager.GetWorkflowInteropService();
			op.Context.Load(instanceService);
			op.Context.Load(interOpService);
			op.Execute();

			Dictionary<string, object> data = new Dictionary<string, object>();

			var allData = new MyFields();
			allData.DataFields.SharePointListItem_RW.Reviewers.Assignment = new Assignment()
			{
				Assignee = new Assignee()
				{
					Person = new List<Person>()
					{
						new Person()
						{
							DisplayName = "Ondrej Valenta",
							AccountId = "i:0#.f|membership|ondrej.valenta@keenmate.com",
							AccountType = "User"
						}
					}
				}
			};
			
			data.Add("myFields", allData);

			IDictionary<string, object> data1 = new Dictionary<string, object>();
			//data1.Add("Reviewers", allData.DataFields.SharePointListItem_RW);
			data1.Add("Assignee", "km\\ondrej.valenta");

			var wfResult = interOpService.StartWorkflow("Collect feedback", Guid.NewGuid(), op.LastList.Id, itemId, data1);
			op.Execute();
			var instancesForListItem = instanceService.EnumerateInstancesForListItem(op.LastList.Id, 2);
			

			//var wa = op.Context.Web.Lists.GetByTitle("xxx").WorkflowAssociations[0];
			//wa.
			//	var wfServicesManager = new WorkflowServicesManager(op.Context, op.LastWeb);
			//InteropService interopService = wfServicesManager.GetWorkflowInteropService();

			//ClientResult<Guid> resultGuid = interopService.StartWorkflow(association.Name, new Guid(), list.Id, itemId, initData);
			//ctx.ExecuteQuery();



			//new WorkflowServicesManager().GetWorkflowInstanceService().
		}

		private static void CreateTermSetAndTerm()
		{
			var op = "http://cpas.shp.cz.keenmate.com"
				.Create(new ConsoleLogger())
				.SetupContext(context =>
				{
					context.Credentials =
						new NetworkCredential() { Domain = ClientSecrets.Domain, Password = ClientSecrets.Password, UserName = ClientSecrets.Username };
				});
			//.SetOnlineCredentials(ClientSecrets.Username, ClientSecrets.Password);
			op.Fail((operation, exception) =>
			{
				Console.WriteLine(exception.Message);
				return operation;
			});

			var customProperties = new Dictionary<string, string>();
			customProperties.Add("Parent", "Parent Term;Parent Term 1");
			customProperties.Add("Custom property 1", $"{DateTime.Now:G}");

			op.OpenTaxonomySession()
				.SelectTaxonomyStore()
				.LoadTermGroup("Term group 2")
				.CreateTermSet("Term group 2 - Term set 1")
				//.CreateTermGroup("Term group 2")
				//.CreateTermSet("Term set")
				//.CreateTerm("Parent term")
				.CreateTerm($"Term - {DateTime.Now:s}", customProperties: customProperties, localProperties: customProperties)
				//.SelectTermGroup("Term group 1")
				//.CreateTermSet("Term group 1 - Term set 2")
				.Execute();
		}
	}


	
}
