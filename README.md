# fluently-sharepoint
CSOM Linq styled helper 


## Create and execute
```
var op = SiteUrl
	.Create(logger)
	.SetOnlineCredentials(UserName, Password) // Available also with SecureString parameter
	.Execute();
```     
## Create, execute and reuse
```
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

```
## Reuse existing CSOM client context
```
ClientContext context = new ClientContext(SiteUrl);
context.Credentials = new SharePointOnlineCredentials(UserName, Password.ToSecureString());

var listTitle = "Documents";

var items = context
	.Create()
	.LoadList(listTitle)
	.GetItems();

logger.Info($"Total items of list {listTitle} with list.ItemCount: {items.Count}");
```			
