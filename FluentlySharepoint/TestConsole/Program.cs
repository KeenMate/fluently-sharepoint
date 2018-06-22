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
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WorkflowServices;
using TestConsole.Loggers;

namespace TestConsole
{
	class Program
	{
		public const string UserName = "trent.reznor@keenmate.com";
		public const string Password = "Discipline042008";
		public const string SiteUrl =
				"https://keenmate.sharepoint.com/sites/demo/fluently-sharepoint/";

		public static ILogger logger = new ConsoleLogger();

		static void Main(string[] args)
		{
			//MeasuredOperation(CreateAndExecute);

			//MeasuredOperation(CreateExecuteAndReuse);

			//MeasuredOperation(ReuseExistingContext);

			//MeasuredOperation(CreateWebAndSeveralListInIt);

			MeasuredOperation(StartStandardWorkflow);
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
				.SetOnlineCredentials(UserName, Password) // Available also with SecureString parameter
				.Execute();

			logger.Trace($"Default timeout: {op.Context.RequestTimeout}");
		}

		private static void CreateExecuteAndReuse()
		{
			logger.Info("Create, execute and reuse example");

			var op = SiteUrl
				.Create(logger)
				.SetOnlineCredentials(UserName, Password) // Available also with SecureString parameter
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
			context.Credentials = new SharePointOnlineCredentials(UserName, Password.ToSecureString());

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
				.SetOnlineCredentials(UserName, Password)
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
		/// Not working. We cannot find out how to send recepients to the standard workflow for it to work
		/// </summary>
		private static void StartStandardWorkflow()
		{
			var op = "http://dev-sp2016-01:7100/"
				.Create(new ConsoleLogger())
				.SetupContext(context =>
				{
					context.Credentials =
						new NetworkCredential() { Domain = "KM", Password = "3.18Fuchsie", UserName = "ondrej.valenta" };
				});
			//.SetOnlineCredentials(UserName, Password);
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


	}


	[XmlRoot(ElementName = "Person", Namespace = "http://schemas.microsoft.com/office/infopath/2007/PartnerControls")]
	public class Person
	{
		[XmlElement(ElementName = "DisplayName", Namespace = "http://schemas.microsoft.com/office/infopath/2007/PartnerControls")]
		public string DisplayName { get; set; }
		[XmlElement(ElementName = "AccountId", Namespace = "http://schemas.microsoft.com/office/infopath/2007/PartnerControls")]
		public string AccountId { get; set; }
		[XmlElement(ElementName = "AccountType", Namespace = "http://schemas.microsoft.com/office/infopath/2007/PartnerControls")]
		public string AccountType { get; set; }
	}

	[XmlRoot(ElementName = "Assignee", Namespace = "http://schemas.microsoft.com/office/infopath/2009/WSSList/dataFields")]
	public class Assignee
	{
		[XmlElement(ElementName = "Person", Namespace = "http://schemas.microsoft.com/office/infopath/2007/PartnerControls")]
		public List<Person> Person { get; set; }
	}

	[XmlRoot(ElementName = "Stage", Namespace = "http://schemas.microsoft.com/office/infopath/2009/WSSList/dataFields")]
	public class Stage
	{
		[XmlAttribute(AttributeName = "nil", Namespace = "http://www.w3.org/2001/XMLSchema-instance")]
		public string Nil { get; set; }
	}

	[XmlRoot(ElementName = "Assignment", Namespace = "http://schemas.microsoft.com/office/infopath/2009/WSSList/dataFields")]
	public class Assignment
	{
		[XmlElement(ElementName = "Assignee", Namespace = "http://schemas.microsoft.com/office/infopath/2009/WSSList/dataFields")]
		public Assignee Assignee { get; set; }
		[XmlElement(ElementName = "Stage", Namespace = "http://schemas.microsoft.com/office/infopath/2009/WSSList/dataFields")]
		public Stage Stage { get; set; }
		[XmlElement(ElementName = "AssignmentType", Namespace = "http://schemas.microsoft.com/office/infopath/2009/WSSList/dataFields")]
		public string AssignmentType { get; set; }
	}

	[XmlRoot(ElementName = "Reviewers", Namespace = "http://schemas.microsoft.com/office/infopath/2009/WSSList/dataFields")]
	public class Reviewers
	{
		[XmlElement(ElementName = "Assignment", Namespace = "http://schemas.microsoft.com/office/infopath/2009/WSSList/dataFields")]
		public Assignment Assignment { get; set; }
	}

	[XmlRoot(ElementName = "DueDateforAllTasks", Namespace = "http://schemas.microsoft.com/office/infopath/2009/WSSList/dataFields")]
	public class DueDateforAllTasks
	{
		[XmlAttribute(AttributeName = "nil", Namespace = "http://www.w3.org/2001/XMLSchema-instance")]
		public string Nil { get; set; }
	}

	[XmlRoot(ElementName = "CC", Namespace = "http://schemas.microsoft.com/office/infopath/2009/WSSList/dataFields")]
	public class CC
	{
		[XmlElement(ElementName = "Person", Namespace = "http://schemas.microsoft.com/office/infopath/2007/PartnerControls")]
		public Person Person { get; set; }
	}

	[XmlRoot(ElementName = "SharePointListItem_RW", Namespace = "http://schemas.microsoft.com/office/infopath/2009/WSSList/dataFields")]
	public class SharePointListItem_RW
	{
		[XmlElement(ElementName = "Reviewers",
			Namespace = "http://schemas.microsoft.com/office/infopath/2009/WSSList/dataFields")]
		public Reviewers Reviewers { get; set; } = new Reviewers();
		[XmlElement(ElementName = "ExpandGroups", Namespace = "http://schemas.microsoft.com/office/infopath/2009/WSSList/dataFields")]
		public string ExpandGroups { get; set; }
		[XmlElement(ElementName = "NotificationMessage", Namespace = "http://schemas.microsoft.com/office/infopath/2009/WSSList/dataFields")]
		public string NotificationMessage { get; set; }
		[XmlElement(ElementName = "DueDateforAllTasks", Namespace = "http://schemas.microsoft.com/office/infopath/2009/WSSList/dataFields")]
		public DueDateforAllTasks DueDateforAllTasks { get; set; } = new DueDateforAllTasks();
		[XmlElement(ElementName = "DurationforSerialTasks", Namespace = "http://schemas.microsoft.com/office/infopath/2009/WSSList/dataFields")]
		public string DurationforSerialTasks { get; set; }
		[XmlElement(ElementName = "DurationUnits", Namespace = "http://schemas.microsoft.com/office/infopath/2009/WSSList/dataFields")]
		public string DurationUnits { get; set; }
		[XmlElement(ElementName = "CC", Namespace = "http://schemas.microsoft.com/office/infopath/2009/WSSList/dataFields")]
		public CC CC { get; set; } = new CC();
		[XmlElement(ElementName = "CancelonChange", Namespace = "http://schemas.microsoft.com/office/infopath/2009/WSSList/dataFields")]
		public string CancelonChange { get; set; }
	}

	[XmlRoot(ElementName = "dataFields", Namespace = "http://schemas.microsoft.com/office/infopath/2003/dataFormSolution")]
	public class DataFields
	{
		[XmlElement(ElementName = "SharePointListItem_RW", Namespace = "http://schemas.microsoft.com/office/infopath/2009/WSSList/dataFields")]
		public SharePointListItem_RW SharePointListItem_RW { get; set; } = new SharePointListItem_RW();
	}

	[XmlRoot(ElementName = "myFields", Namespace = "http://schemas.microsoft.com/office/infopath/2003/dataFormSolution")]
	public class MyFields
	{
		[XmlElement(ElementName = "queryFields", Namespace = "http://schemas.microsoft.com/office/infopath/2003/dataFormSolution")]
		public string QueryFields { get; set; }
		[XmlElement(ElementName = "dataFields", Namespace = "http://schemas.microsoft.com/office/infopath/2003/dataFormSolution")]
		public DataFields DataFields { get; set; } = new DataFields();
		[XmlAttribute(AttributeName = "xsd", Namespace = "http://www.w3.org/2000/xmlns/")]
		public string Xsd { get; set; }
		[XmlAttribute(AttributeName = "dms", Namespace = "http://www.w3.org/2000/xmlns/")]
		public string Dms { get; set; }
		[XmlAttribute(AttributeName = "dfs", Namespace = "http://www.w3.org/2000/xmlns/")]
		public string Dfs { get; set; }
		[XmlAttribute(AttributeName = "q", Namespace = "http://www.w3.org/2000/xmlns/")]
		public string Q { get; set; }
		[XmlAttribute(AttributeName = "d", Namespace = "http://www.w3.org/2000/xmlns/")]
		public string D { get; set; }
		[XmlAttribute(AttributeName = "ma", Namespace = "http://www.w3.org/2000/xmlns/")]
		public string Ma { get; set; }
		[XmlAttribute(AttributeName = "pc", Namespace = "http://www.w3.org/2000/xmlns/")]
		public string Pc { get; set; }
		[XmlAttribute(AttributeName = "xsi", Namespace = "http://www.w3.org/2000/xmlns/")]
		public string Xsi { get; set; }
	}
}
