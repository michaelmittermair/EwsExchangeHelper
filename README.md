# EwsExchangeHelper
Functions to connect to exchange with EWS (Exchange Web Services)

## Infos
This repo offers you a package to connect to exchange web services. You can get items from exchange, especially contacts, groups and appointments.


## References
* Some of the codeparts I used are from the examples of https://github.com/sherlock1982/ews-managed-api
* Additional information and code examples used from https://code.msdn.microsoft.com/Exchange-2013-101-Code-3c38582c

## Example
### using the Service to get all tasks of an user

```
using (var service = new MsExchangeServices())
{
	service.ImpersonateUser(ConnectingIdType.SmtpAddress, "max.mustermann@outlook.com");

	var items = service.GetItems(WellKnownFolderName.Tasks, new ItemView(100));
	
		foreach (var item in items.Items)
		{
			//item.Id;
			//item.Subject;
			//item.DateTimeCreated;
			//item.DateTimeReceived;
			//item.DateTimeSent;
			//item.LastModifiedName;
			//item.LastModifiedTime;
			//item.ParentFolderId;
		}
	}
}
```

### using the service to add permissions to folders
With this function, you can add permissions like Read-Permission on the Inbox

```
using (var service = new MsExchangeServices() {Logger = _log })
{
	service.ImpersonateUser(ConnectingIdType.SmtpAddress, "max.mustermann@outlook.com");

	EnableFolderPermissions(service.ExchangeService, "mina.mustermann@outlook.com", WellKnownFolderName.Root);
	EnableFolderPermissions(service.ExchangeService, "mina.mustermann@outlook.com", WellKnownFolderName.Inbox);
}
```