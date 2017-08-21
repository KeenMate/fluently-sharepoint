using Microsoft.SharePoint.Client;

namespace FluentlySharepoint
{
	public class DeferredAction
	{
		public ClientObject ClientObject { get; set; }
		public DeferredActions Action { get; set; }
	}
}