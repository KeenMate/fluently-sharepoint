using System.Collections.Generic;
using System.Xml.Serialization;

namespace TestConsole.WorkflowModels
{
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