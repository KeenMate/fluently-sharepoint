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
