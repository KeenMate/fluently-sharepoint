using Microsoft.SharePoint.Client;

namespace KeenMate.FluentlySharePoint
{
	public class DeferredAction
	{
		public ClientObject ClientObject { get; set; }
		public DeferredActions Action { get; set; }
	}
}